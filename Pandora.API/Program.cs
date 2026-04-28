using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.RateLimiting;
using FluentValidation;
using FluentValidation.AspNetCore;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.Cors.Infrastructure;
using Microsoft.AspNetCore.RateLimiting;
using Microsoft.AspNetCore.StaticFiles;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.FileProviders;
using Microsoft.IdentityModel.Tokens;
using Microsoft.OpenApi.Models;
using Pandora.API.Hubs;
using Pandora.API.Middleware;
using Pandora.API.Services;
using Pandora.Application.Features.Campaigns;
using Pandora.Application.Features.Users;
using Pandora.Application.Interfaces;
using Pandora.Application.Validators;
using Pandora.Domain.Entities;
using Pandora.Domain.Enums;
using Pandora.Infrastructure;
using Pandora.Infrastructure.Persistence;
using Pandora.Infrastructure.Services;
using Swashbuckle.AspNetCore.SwaggerGen;
using Swashbuckle.AspNetCore.SwaggerUI;

var builder = WebApplication.CreateBuilder(args);

// ── Resolver pipe de LocalDB si aplica ──────────────────────────────────────
string rawConnStr = builder.Configuration.GetConnectionString("PandoraDb") ?? "";
string resolvedConnStr = ResolveLocalDbPipe(rawConnStr);
if (resolvedConnStr != rawConnStr)
    builder.Configuration["ConnectionStrings:PandoraDb"] = resolvedConnStr;

// ── MVC + FluentValidation ───────────────────────────────────────────────────
builder.Services.AddControllers()
    .AddJsonOptions(opts =>
    {
        opts.JsonSerializerOptions.NumberHandling =
            System.Text.Json.Serialization.JsonNumberHandling.AllowReadingFromString;
    });
builder.Services.AddFluentValidationAutoValidation();
builder.Services.AddValidatorsFromAssemblyContaining<CreateCampaignRequestValidator>();
builder.Services.AddEndpointsApiExplorer();

// ── Swagger ──────────────────────────────────────────────────────────────────
builder.Services.AddSwaggerGen(c =>
{
    c.SwaggerDoc("v1", new OpenApiInfo
    {
        Title       = "Pandora API",
        Version     = "v1",
        Description = "Sistema de envio masivo de correos - iMET"
    });
    c.AddSecurityDefinition("Bearer", new OpenApiSecurityScheme
    {
        Name        = "Authorization",
        Type        = SecuritySchemeType.Http,
        Scheme      = "bearer",
        BearerFormat = "JWT",
        In          = ParameterLocation.Header
    });
    c.AddSecurityRequirement(new OpenApiSecurityRequirement
    {
        {
            new OpenApiSecurityScheme
            {
                Reference = new OpenApiReference { Type = ReferenceType.SecurityScheme, Id = "Bearer" }
            },
            Array.Empty<string>()
        }
    });
});

// ── CORS ─────────────────────────────────────────────────────────────────────
string[] allowedOrigins = builder.Configuration.GetSection("AllowedOrigins").Get<string[]>()
                          ?? ["http://localhost:3000"];
builder.Services.AddCors(opt =>
{
    opt.AddPolicy("PandoraPolicy", p =>
        p.WithOrigins(allowedOrigins)
         .AllowAnyHeader()
         .AllowAnyMethod()
         .AllowCredentials());
});

// ── JWT Authentication ───────────────────────────────────────────────────────
string jwtKey = builder.Configuration["JwtSettings:Key"]!;
builder.Services.AddAuthentication("Bearer").AddJwtBearer(opts =>
{
    opts.TokenValidationParameters = new TokenValidationParameters
    {
        ValidateIssuer           = true,
        ValidateAudience         = true,
        ValidateLifetime         = true,
        ValidateIssuerSigningKey = true,
        ValidIssuer              = builder.Configuration["JwtSettings:Issuer"],
        ValidAudience            = builder.Configuration["JwtSettings:Audience"],
        IssuerSigningKey         = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(jwtKey))
    };
    opts.Events = new JwtBearerEvents
    {
        OnMessageReceived = context =>
        {
            var accessToken = context.Request.Query["access_token"];
            if (!string.IsNullOrEmpty(accessToken))
            {
                var path = context.HttpContext.Request.Path;
                if (path.StartsWithSegments("/hubs") ||
                    path.StartsWithSegments("/api/inventory/excel") ||
                    (path.Value?.Contains("/visualizar") == true))
                {
                    context.Token = accessToken;
                }
            }
            return Task.CompletedTask;
        }
    };
});

// ── SignalR ──────────────────────────────────────────────────────────────────
builder.Services.AddSignalR();
builder.Services.AddScoped<IProgressNotifier, SignalRProgressNotifier>();

// ── Infraestructura (EF Core, repositorios, servicios) ───────────────────────
builder.Services.AddInfrastructure(builder.Configuration);

if (builder.Environment.IsDevelopment())
    builder.Services.AddHostedService<LocalDbKeepAliveService>();

// ── Rate Limiting — protege /auth/login contra fuerza bruta ─────────────────
builder.Services.AddRateLimiter(options =>
{
    // Ventana fija: máx 5 intentos de login por IP por minuto
    options.AddFixedWindowLimiter("login-policy", config =>
    {
        config.PermitLimit          = 5;
        config.Window               = TimeSpan.FromMinutes(1);
        config.QueueProcessingOrder = QueueProcessingOrder.OldestFirst;
        config.QueueLimit           = 0;
    });
    // Tickets: máx 10 creaciones por IP cada 5 minutos (anti-spam)
    options.AddFixedWindowLimiter("tickets-policy", config =>
    {
        config.PermitLimit          = 10;
        config.Window               = TimeSpan.FromMinutes(5);
        config.QueueProcessingOrder = QueueProcessingOrder.OldestFirst;
        config.QueueLimit           = 0;
    });
    options.RejectionStatusCode = StatusCodes.Status429TooManyRequests;
    options.OnRejected = async (context, token) =>
    {
        context.HttpContext.Response.ContentType = "application/json";
        await context.HttpContext.Response.WriteAsync(
            "{\"error\":\"Demasiados intentos. Espera 1 minuto e intenta de nuevo.\"}",
            token);
    };
});

// ════════════════════════════════════════════════════════════════════════════
var app = builder.Build();
// ════════════════════════════════════════════════════════════════════════════

// ── Security Headers (Sprint 2) ───────────────────────────────────────────
app.UseSecurityHeaders();

// ── Rate Limiter (Sprint 2) ───────────────────────────────────────────────
app.UseRateLimiter();

// ── Directorios de almacenamiento ─────────────────────────────────────────
string wwwroot     = app.Environment.WebRootPath ?? Path.Combine(Directory.GetCurrentDirectory(), "wwwroot");
Directory.CreateDirectory(Path.Combine(wwwroot,   "uploads", "profiles"));
Directory.CreateDirectory(Path.Combine(wwwroot,   "uploads", "banners"));
Directory.CreateDirectory(Path.Combine(wwwroot,   "uploads", "templates"));
string storageRoot = Path.Combine(app.Environment.ContentRootPath, "storage");
Directory.CreateDirectory(Path.Combine(storageRoot, "libros"));
Directory.CreateDirectory(Path.Combine(storageRoot, "portadas"));

// ── Swagger ───────────────────────────────────────────────────────────────
app.UseSwagger();
app.UseSwaggerUI(c => c.SwaggerEndpoint("/swagger/v1/swagger.json", "Pandora v1"));

// ── Archivos estáticos ────────────────────────────────────────────────────
app.UseDefaultFiles();
app.UseStaticFiles();
string portadasPath = Path.Combine(app.Environment.ContentRootPath, "storage", "portadas");
Directory.CreateDirectory(portadasPath);
app.UseStaticFiles(new StaticFileOptions
{
    FileProvider        = new PhysicalFileProvider(portadasPath),
    RequestPath         = "/api/storage/portadas",
    ContentTypeProvider = new FileExtensionContentTypeProvider()
});

// ── Pipeline ──────────────────────────────────────────────────────────────
app.UseCors("PandoraPolicy");
app.UseAuthentication();
app.UseAuthorization();
app.MapControllers();
app.MapHub<ProgressHub>("/hubs/progress");
app.MapFallbackToFile("index.html");

// ── Migraciones + Seed ────────────────────────────────────────────────────
using (var scope = app.Services.CreateScope())
{
    var db = scope.ServiceProvider.GetRequiredService<PandoraDbContext>();
    await db.Database.MigrateAsync();

    // ── Departamentos ─────────────────────────────────────────────────────
    (string Name, string Prefix)[] deptData =
    [
        ("Tecnologías de Información",       "EQT-TI"),
        ("Control Escolar",                  "EQT-CE"),
        ("Dirección de Programas Académicos","EQT-DPA"),
        ("Dirección de Administración",      "EQT-DA"),
        ("Dirección General",                "EQT-DG"),
        ("Asistente de Dirección",           "EQT-AD"),
        ("Coordinación de Mercadotecnia",    "EQT-MKT"),
        ("Dirección Comercial",              "EQT-DCOM"),
        ("Coordinación Comercial",           "EQT-COM"),
    ];
    bool deptChanged = false;
    foreach (var (name, prefix) in deptData)
    {
        var existing = db.Departments.FirstOrDefault(d => d.Name == name);
        if (existing == null)
        {
            db.Departments.Add(new Department
            {
                Name = name, InventoryPrefix = prefix,
                IsActive = true, CreatedAt = DateTime.UtcNow
            });
            deptChanged = true;
        }
        else if (existing.InventoryPrefix != prefix)
        {
            existing.InventoryPrefix = prefix;
            deptChanged = true;
        }
    }
    if (deptChanged) await db.SaveChangesAsync();

    // ── Salas ─────────────────────────────────────────────────────────────
    (string Name, int Cap, string Loc, string Color)[] roomData =
    [
        ("AULA 1",            20, "Edificio Principal", "#1976d2"),
        ("AULA 2",            20, "Edificio Principal", "#388e3c"),
        ("AULA 3",            20, "Edificio Principal", "#d32f2f"),
        ("AULA 4",            20, "Edificio Principal", "#7b1fa2"),
        ("AULA 5",            20, "Edificio Principal", "#f57c00"),
        ("CENTRO DE COMPUTO", 30, "Edificio TI",        "#0288d1"),
        ("CABINA DE PODCAST",  4, "Edificio Principal", "#455a64"),
    ];
    bool roomChanged = false;
    foreach (var (name, cap, loc, color) in roomData)
    {
        if (!db.Rooms.Any(r => r.Name == name))
        {
            db.Rooms.Add(new Room
            {
                Name = name, Capacity = cap, Location = loc, Color = color,
                IsActive = true, CreatedAt = DateTime.UtcNow
            });
            roomChanged = true;
        }
    }
    if (roomChanged) await db.SaveChangesAsync();

    // ── Admin inicial ─────────────────────────────────────────────────────
    if (!db.AppUsers.Any())
    {
        string adminUsername = builder.Configuration["AdminUser:Username"] ?? "admin";
        string adminPassword = builder.Configuration["AdminUser:Password"] ?? "PandoraAdmin2024!";
        db.AppUsers.Add(new AppUser
        {
            Username     = adminUsername.ToLower(),
            FullName     = "Administrador iMET",
            Email        = "sistemas@imet.edu.mx",
            PasswordHash = UserService.HashPassword(adminPassword),
            Role         = "Admin",
            Modules      = ModuleAccess.All,
            IsActive     = true
        });
        await db.SaveChangesAsync();
    }

    // ── Biblioteca — seed inicial de categorías y libros ─────────────────
    if (!db.Categorias.Any())
    {
        var now = DateTime.UtcNow;
        (string Nombre, string Icono, string Nivel,
         (string Archivo, long Tamano, string Titulo, string Autor, int? Anio)[] Libros)[] catalogoSeed =
        [
            ("Lengua y Comunicación", "📚", "Primer Semestre",
            [
                ("LENGUA-Y-COMUNICACION.pdf",                                          1_846_083, "Lengua y Comunicación",                        "DGB iMET", null),
                ("Lengua Comunicacion y Cultura Digital I - SEP DGB 2024.pdf",         6_232_725, "Lengua, Comunicación y Cultura Digital I",      "SEP DGB",  2024),
                ("Lengua y comunicación I.pdf",                                          995_005, "Lengua y Comunicación I",                       "SEP",      null),
            ]),
            ("Pensamiento Matemático", "📐", "Primer Semestre",
            [
                ("Pensamiento Matematico y Tecnologia - SEP DGB 2024.pdf",            57_423_800, "Pensamiento Matemático y Tecnología",           "SEP DGB",  2024),
                ("Pensamiento Matemático 1.pdf",                                        2_162_750, "Pensamiento Matemático 1",                      "SEP",      null),
                ("Pensamiento Matemático I Cuevas Martínez, Betsy; Rivera Castillo, Enrique 2024.pdf", 968_700, "Pensamiento Matemático I", "Cuevas Martínez, Betsy; Rivera Castillo, Enrique", 2024),
            ]),
            ("Ciencias Naturales Experimentales y Tecnología", "🔬", "Segundo Semestre",
            [
                ("Ciencias Naturales 1 Experiencias y Tecnología - Adrian Lecona - 2024.pdf", 2_158_428, "Ciencias Naturales 1: Experiencias y Tecnología", "Lecona, Adrián", 2024),
                ("Ciencias Naturales, Experimentales y Tecnología I - Reyes, Ana Karen; Hernández, Angélica -2024.pdf", 2_158_428, "Ciencias Naturales, Experimentales y Tecnología I", "Reyes, Ana Karen; Hernández, Angélica", 2024),
                ("La Materia y sus Interacciones - DGB 2023.pdf", 957_631, "La Materia y sus Interacciones", "DGB", 2023),
            ]),
            ("Taller de Ciencias", "🧪", "Segundo Semestre",
            [
                ("Taller de Ciencias 1 - Morales Ángel, Elsa Ivonne, Méndez Rentería, Danelia - 2024.pdf", 501_992, "Taller de Ciencias 1", "Morales Ángel, Elsa Ivonne; Méndez Rentería, Danelia", 2024),
                ("Taller de Ciencias I - Programa de Estudios DGB 2024.pdf", 501_992, "Taller de Ciencias I - Programa de Estudios", "DGB", 2024),
                ("Taller de Ciencias II - Programa de Estudios DGB 2024.pdf", 889_727, "Taller de Ciencias II - Programa de Estudios", "DGB", 2024),
            ]),
            ("Ciencias Sociales", "🌎", "Primer Semestre",
            [
                ("Ciencias Sociales I - Programa de Estudios DGB 2023.pdf", 1_144_382, "Ciencias Sociales I - Programa de Estudios", "DGB", 2023),
                ("Introduccion a las Ciencias Sociales - SEP DGB 2024.pdf", 19_820_418, "Introducción a las Ciencias Sociales", "SEP DGB", 2024),
                ("Introducción a las Ciencias Sociales - Tovar González, Rafael Manuel - 2024.pdf", 19_820_418, "Introducción a las Ciencias Sociales", "Tovar González, Rafael Manuel", 2024),
            ]),
            ("Laboratorio de Investigación", "🔍", "Tercer Semestre",
            [
                ("Laboratorio de Investigacion - Programa de Estudios DGB 2024.pdf", 839_885, "Laboratorio de Investigación - Programa de Estudios", "DGB", 2024),
                ("Laboratorio de Investigación - Medina Gual, Luis - 2024.pdf", 708_093, "Laboratorio de Investigación", "Medina Gual, Luis", 2024),
                ("Metodologia de la Investigacion - SEP Telebachillerato 2024.pdf", 33_490_731, "Metodología de la Investigación", "SEP Telebachillerato", 2024),
            ]),
            ("Conciencia Histórica", "📜", "Segundo Semestre",
            [
                ("Conciencia Historica I - Programa de Estudios DGB 2025.pdf", 867_185, "Conciencia Histórica I - Programa de Estudios", "DGB", 2025),
                ("Conciencia Histórica 1  Serie Perspectivas.pdf", 1_768_068, "Conciencia Histórica 1: Serie Perspectivas", "SEP", null),
                ("Historia de Mexico I - SEP Telebachillerato 2024.pdf", 23_204_582, "Historia de México I", "SEP Telebachillerato", 2024),
            ]),
            ("Pensamiento Filosófico y Humanidades", "💭", "Cuarto Semestre",
            [
                ("Pensamiento Filosófico y Humanidades 1 - Escobar Valenzuela, Gustavo Alberto - 2024.pdf", 1_077_388, "Pensamiento Filosófico y Humanidades 1", "Escobar Valenzuela, Gustavo Alberto", 2024),
                ("Pensamiento Filosófico y Humanidades 1 Propósitos Formativos.pdf", 1_920_123, "Pensamiento Filosófico y Humanidades 1: Propósitos Formativos", "SEP DGB", null),
                ("Pensamiento Filosófico y Humanidades I - Gómez Navas Chapa, Leonardo - 2025.pdf", 1_920_123, "Pensamiento Filosófico y Humanidades I", "Gómez Navas Chapa, Leonardo", 2025),
            ]),
            ("Pensamiento Literario", "✍️", "Tercer Semestre",
            [
                ("Literatura II - SEP Telebachillerato 2019.pdf", 14_169, "Literatura II", "SEP Telebachillerato", 2019),
                ("Pensamiento Literario Ávila Vázquez, Leonardo Mauricio 2024.pdf", 733_030, "Pensamiento Literario", "Ávila Vázquez, Leonardo Mauricio", 2024),
                ("Pensamiento Literario Márquez Escamilla, Aldo Daniel 2024.pdf", 733_030, "Pensamiento Literario", "Márquez Escamilla, Aldo Daniel", 2024),
            ]),
            ("Cultura Digital", "💻", "Primer Semestre",
            [
                ("Cultura Digital I - Programa de Estudios DGB 2023.pdf", 882_428, "Cultura Digital I - Programa de Estudios", "DGB", 2023),
                ("Cultura Digital II - Programa de Estudios DGB 2025.pdf", 891_340, "Cultura Digital II - Programa de Estudios", "DGB", 2025),
                ("Ramírez Martinell, A. Háblame de TIC. Tecnología digital en la educación superior. (Tomo 1).pdf", 822_801, "Háblame de TIC: Tecnología digital en la educación superior", "Ramírez Martinell, A.", null),
            ]),
            ("Inglés", "🇬🇧", "Primer Semestre",
            [
                ("Essential-english-for-foreign-students-book-1.pdf", 7_696_289, "Essential English for Foreign Students - Book 1", "Eckersley, C.E.", null),
                ("Lengua Extranjera Ingles I - Programa de Estudios DGB 2023.pdf", 993_286, "Lengua Extranjera: Inglés I - Programa de Estudios", "DGB", 2023),
                ("Oral-Communication-for-Non-Native-Speakers-of-English-2nd-Edition-1742218240.pdf", 1_182_339, "Oral Communication for Non-Native Speakers of English", "Varios", null),
            ]),
            ("Formación Socioemocional", "❤️", "General",
            [
                ("Didactica para la Formacion Socioemocional - DGB 2024.pdf", 4_137_394, "Didáctica para la Formación Socioemocional", "DGB", 2024),
                ("Educación Socioemocional (NEM) - Burguette - Calderón - 2023.pdf", 2_383_807, "Educación Socioemocional (NEM)", "Burguette; Calderón", 2023),
                ("Recursos Socioemocionales - Programa DGB 2023.pdf", 1_238_470, "Recursos Socioemocionales - Programa de Estudios", "DGB", 2023),
            ]),
            ("Fundamentos de Administración", "🏢", "Tercer Semestre",
            [
                ("Administracion 13a edición - Robbins Stephen P..pdf", 17_163_546, "Administración (13a edición)", "Robbins, Stephen P.", null),
                ("Introduccion a la Teoria General de la Administracion - Chiavenato.pdf", 96_285_461, "Introducción a la Teoría General de la Administración", "Chiavenato, Idalberto", null),
                ("Koontz-Harold-Administracion-Una-Perspectiva-Global.pdf", 22_120_287, "Administración: Una Perspectiva Global", "Koontz, Harold", null),
            ]),
            ("Procesos Contables", "📊", "Cuarto Semestre",
            [
                ("2009_contabilidad-financiera-i.pdf", 2_569_501, "Contabilidad Financiera I", "SEP", 2009),
                ("2011_principios-de-contabilidad_una-perspectiva-empresarial_vol-2.pdf", 2_859_626, "Principios de Contabilidad: Una Perspectiva Empresarial", "Varios", 2011),
                ("Principios de Contabilidad 4a ed - Romero Lopez - U Veracruzana.pdf", 4_178_540, "Principios de Contabilidad (4a ed.)", "Romero López, Javier", null),
            ]),
            ("Administración", "📋", "Quinto Semestre",
            [
                ("Administracion_Stephen_P_Robbins..pdf", 17_465_998, "Administración", "Robbins, Stephen P.", null),
                ("Administracion_Zacarias torres hernandez.pdf", 14_056_863, "Administración", "Torres Hernández, Zacarías", null),
                ("Administración - 10ma Edición - Stephen P. Robbins & Mary Coulter (1).pdf", 173_679_924, "Administración (10ma Edición)", "Robbins, S. P.; Coulter, Mary", null),
            ]),
            ("Contabilidad", "💰", "Quinto Semestre",
            [
                ("Administracion 13a edición - Robbins Stephen P. 2004.pdf", 3_221_165, "Administración (13a edición, 2004)", "Robbins, Stephen P.", 2004),
                ("Principios de Contabilidad - Romero Lopez - U Veracruzana.pdf", 4_178_540, "Principios de Contabilidad", "Romero López, Javier", null),
                ("PrincipiosdeContabilidad_Alvaro Javier Romero Lopez 4ta Edicion.pdf", 4_178_540, "Principios de Contabilidad (4ta Ed.)", "Romero López, Álvaro Javier", null),
            ]),
            ("Derecho y Sociedad", "⚖️", "Cuarto Semestre",
            [
                ("Derecho y Sociedad I - Programa de Estudios DGB 2025.pdf", 850_936, "Derecho y Sociedad I - Programa de Estudios", "DGB", 2025),
                ("Introducción al Estudio del Derecho García Máynez, Eduardo 2020.pdf", 6_091_846, "Introducción al Estudio del Derecho", "García Máynez, Eduardo", 2020),
                ("Introducción al Estudio del Derecho – Pereznieto CastroDerecho-ciencias-sociales-2014-2015.pdf", 61_970_104, "Introducción al Estudio del Derecho", "Pereznieto Castro", null),
            ]),
            ("Cultura de la Legalidad y Ciudadanía", "🏛️", "Tercer Semestre",
            [
                ("Cultura-de-la-legalidad-y-derechos-humanos-Jhonatan Alejandro Correa Ortiz.pdf", 1_295_297, "Cultura de la Legalidad y Derechos Humanos", "Correa Ortiz, Jhonatan Alejandro", null),
                ("Democracia y Cultura de la Legalidad - Pedro Salazar Ugarte - UNAM.pdf", 302_433, "Democracia y Cultura de la Legalidad", "Salazar Ugarte, Pedro", null),
                ("cultura-de-la-legalidad-e-instituciones-en-mexico- Julia Isabel Flores.pdf", 692_348, "Cultura de la Legalidad e Instituciones en México", "Flores, Julia Isabel", null),
            ]),
            ("Innovación y Tecnologías Emergentes", "🚀", "Sexto Semestre",
            [
                ("Aprender_y_educar_con_las_tecnologias_del_Siglo_XXI.pdf", 18_830_977, "Aprender y Educar con las Tecnologías del Siglo XXI", "Varios", null),
                ("Losada,I. Innovación educativa y formación docente. Últimas aportaciones en la investigación, Isidoro Hernán Losada.pdf", 14_363_681, "Innovación Educativa y Formación Docente", "Losada, Isidoro Hernán", null),
                ("Tecnologias e innovacion en la practica educativa - Ruben Edel Navarro.pdf", 18_518_031, "Tecnologías e Innovación en la Práctica Educativa", "Edel Navarro, Rubén", null),
            ]),
            ("Principios de Programación y Desarrollo de Software", "⌨️", "Quinto Semestre",
            [
                ("Fundamentos de programación en Java - Jorge Martinez Ladrón de Guevara.pdf", 807_190, "Fundamentos de Programación en Java", "Martínez Ladrón de Guevara, Jorge", null),
                ("Fundamentos de programación, 4ta Edición - Luis Joyanes Aguilar-.pdf", 21_897_183, "Fundamentos de Programación (4ta Edición)", "Joyanes Aguilar, Luis", null),
                ("Fundamentos-de-Programacion Luis Joyanes Aguilar.pdf", 3_849_209, "Fundamentos de Programación", "Joyanes Aguilar, Luis", null),
            ]),
            ("Pensamiento Lógico y Matemático Aplicado", "🧮", "Cuarto Semestre",
            [
                ("Fundamentos De Matematicas Para Bachillerato.pdf", 65_134_552, "Fundamentos de Matemáticas para Bachillerato", "Varios", null),
                ("Pensamiento Matematico I - Programa DGB 2023.pdf", 968_700, "Pensamiento Matemático I - Programa de Estudios", "DGB", 2023),
                ("Probabilidad y Estadistica I - Programa DGB 2023.pdf", 356_760, "Probabilidad y Estadística I - Programa de Estudios", "DGB", 2023),
            ]),
            ("Electrónica", "⚡", "Sexto Semestre",
            [
                ("Curso de electrónica básica-Luis Larrion.pdf", 14_599_931, "Curso de Electrónica Básica", "Larrion, Luis", null),
                ("Electronica-Fundamental-7- Jose M. Angulo - Juan Jose Lopez Benlloch.pdf", 45_072_627, "Electrónica Fundamental (7a ed.)", "Angulo, José M.; López Benlloch, Juan José", null),
                ("Libro Electronica basica - Proyecto Descartes - 2022.pdf", 22_773_902, "Electrónica Básica", "Proyecto Descartes", 2022),
            ]),
            ("Intervención en la Educación Obligatoria", "🎓", "Sexto Semestre",
            [
                ("Barraza Macías,A. Elaboración de propuestas de intervención educativa.pdf", 2_337_166, "Elaboración de Propuestas de Intervención Educativa", "Barraza Macías, Arturo", null),
                ("Estrategias Docentes para un Aprendizaje Significativo – Díaz Barriga.pdf", 7_107_116, "Estrategias Docentes para un Aprendizaje Significativo", "Díaz Barriga, Frida", null),
                ("Intervencion en la Educacion Obligatoria - Programa DGB 2024.pdf", 809_514, "Intervención en la Educación Obligatoria", "DGB", 2024),
            ]),
            ("Tecnología de la Información y Comunicación", "🖥️", "Segundo Semestre",
            [
                ("Informatica - Programa DGB 2023.pdf", 354_202, "Informática - Programa de Estudios", "DGB", 2023),
                ("Introduccion a la Computacion - Peter Norton - U Los Andes.pdf", 234_038, "Introducción a la Computación", "Norton, Peter", null),
                ("Tecnologias de la Comunicacion y la Informacion - Programa DGB 2024.pdf", 720_580, "Tecnologías de la Comunicación y la Información", "DGB", 2024),
            ]),
        ];

        foreach (var cat in catalogoSeed)
        {
            var categoria = new Categoria
            {
                Nombre      = cat.Nombre,
                Icono       = cat.Icono,
                Descripcion = $"Material bibliográfico de {cat.Nombre} para Media Superior iMET",
                Activo      = true,
                CreatedAt   = now
            };
            db.Categorias.Add(categoria);
            await db.SaveChangesAsync();

            foreach (var libro in cat.Libros)
            {
                db.Libros.Add(new Libro
                {
                    Titulo          = libro.Titulo,
                    Autor           = libro.Autor,
                    NivelEducativo  = cat.Nivel,
                    Idioma          = "Español",
                    RutaArchivo     = $"libros/{libro.Archivo}",
                    TamanoArchivo   = libro.Tamano,
                    Activo          = true,
                    CategoriaId     = categoria.Id,
                    AnioPublicacion = libro.Anio,
                    CreatedAt       = now
                });
            }
            await db.SaveChangesAsync();
        }
    }
}

// ── Tabla Licencias (fuera del modelo EF — migración manual) ─────────────────
using (var scope2 = app.Services.CreateScope())
{
    var cfg = scope2.ServiceProvider.GetRequiredService<IConfiguration>();
    await using var conn = new Microsoft.Data.SqlClient.SqlConnection(
        cfg.GetConnectionString("PandoraDb"));
    await conn.OpenAsync();

    // Crear tabla si no existe
    await using var cmdCreate = conn.CreateCommand();
    cmdCreate.CommandText = """
        IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'Licencias' AND schema_id = SCHEMA_ID('dbo'))
        BEGIN
            CREATE TABLE dbo.Licencias (
                Id             INT IDENTITY(1,1) PRIMARY KEY,
                Numero         INT           NOT NULL,
                Plataforma     NVARCHAR(100) NOT NULL,
                Area           NVARCHAR(50)  NOT NULL,
                Responsable    NVARCHAR(100) NULL,
                FrecuenciaPago NVARCHAR(20)  NOT NULL,
                FechaInicio    DATE          NOT NULL,
                ProximoPago    DATE          NOT NULL,
                CostoMXN       DECIMAL(12,2) NOT NULL DEFAULT 0,
                Estado         NVARCHAR(20)  NOT NULL DEFAULT 'Activa',
                Notas          NVARCHAR(500) NULL,
                CreadoEn       DATETIME2     NOT NULL DEFAULT GETUTCDATE(),
                ActualizadoEn  DATETIME2     NOT NULL DEFAULT GETUTCDATE()
            );
        END
        """;
    await cmdCreate.ExecuteNonQueryAsync();

    // Seed inicial solo si la tabla está vacía
    await using var cmdCount = conn.CreateCommand();
    cmdCount.CommandText = "SELECT COUNT(*) FROM dbo.Licencias";
    int count = (int)(await cmdCount.ExecuteScalarAsync())!;

    if (count == 0)
    {
        await using var cmdSeed = conn.CreateCommand();
        cmdSeed.CommandText = """
            INSERT INTO dbo.Licencias (Numero,Plataforma,Area,Responsable,FrecuenciaPago,FechaInicio,ProximoPago,CostoMXN,Estado,Notas) VALUES
            (1,'E-Study','TI',NULL,'Mensual','2025-01-01','2026-05-01',5220.00,'Por vencer','Plataforma educativa'),
            (2,'Fortinet FortiGate','TI',NULL,'Anual','2025-07-16','2026-07-16',15000.00,'Activa','Firewall de seguridad'),
            (3,'Google Workspace','TI',NULL,'Mensual','2025-01-01','2026-05-01',0.00,'Por vencer','Administración de correos'),
            (4,'Canva','Marketing',NULL,'Anual','2025-12-07','2026-12-07',1210.00,'Activa','Editor de Imágenes'),
            (5,'Capcut','Marketing',NULL,'Anual','2025-11-14','2026-11-14',2400.00,'Activa','Editor de videos'),
            (6,'OpenAI','Marketing',NULL,'Mensual','2025-11-07','2026-05-07','0.00','Cancelada','IA aplicada a recursos Académicos'),
            (7,'Envato','Marketing',NULL,'Anual','2024-10-16','2026-04-16',0.00,'Cancelada','Administración de diseños'),
            (8,'Adobe Creative','Marketing',NULL,'Anual','2025-03-04','2027-03-06',1055.29,'Activa','Gestor de Diseño'),
            (9,'Claude Max','Marketing',NULL,'Mensual','2025-02-05','2026-05-11',1785.10,'Por vencer','IA aplicada a recursos Académicos'),
            (10,'Claude Max','Innovación',NULL,'Mensual','2026-04-01','2026-05-01',1785.10,'Por vencer','IA aplicada a recursos Académicos'),
            (11,'Claude Basic','Socios',NULL,'Mensual','2026-04-09','2026-05-09',357.00,'Activa','IA aplicada a recursos Académicos'),
            (12,'Claude Basic','TI',NULL,'Mensual','2026-04-01','2026-05-01',357.02,'Por vencer','IA aplicada a recursos Académicos'),
            (13,'Antivirus Kaspersky','TI',NULL,'Anual','2025-09-19','2026-09-19',2000.00,'Activa','Seguridad a equipos de cómputo'),
            (14,'Antivirus Kaspersky','TI',NULL,'Anual','2025-09-21','2026-09-21',2000.00,'Activa','Seguridad a equipos de cómputo'),
            (15,'Antivirus Kaspersky','TI',NULL,'Anual','2025-09-21','2026-09-21',2000.00,'Activa','Seguridad a equipos de cómputo'),
            (16,'Antivirus Kaspersky','TI',NULL,'Anual','2025-09-22','2026-09-22',2000.00,'Activa','Seguridad a equipos de cómputo'),
            (17,'Antivirus Kaspersky','TI',NULL,'Anual','2025-09-22','2026-09-22',2000.00,'Activa','Seguridad a equipos de cómputo'),
            (18,'Antivirus Kaspersky','Marketing',NULL,'Anual','2025-09-25','2026-09-25',2000.00,'Activa','Seguridad a equipos de cómputo'),
            (19,'GitHub','Innovación',NULL,'Mensual','2026-02-09','2026-05-10',264.53,'Por vencer','IA aplicada a Desarrollo de Software'),
            (20,'DigitalOcean','Innovación',NULL,'Mensual','2026-02-01','2026-05-01',496.71,'Por vencer','Servidor Cloud para plataformas'),
            (21,'Akky','TI',NULL,'Anual','2025-02-12','2026-05-06',2000.00,'Por vencer','Alojamiento de la página web'),
            (22,'HiggsField','Marketing',NULL,'Mensual','2026-04-08','2026-05-08',887.40,'Por vencer','Generador de Imágenes'),
            (23,'Midjourney','Marketing',NULL,'Mensual','2026-04-08','2026-05-08',637.78,'Por vencer','Generador de Imágenes');
            """;
        await cmdSeed.ExecuteNonQueryAsync();
    }
}

// ── Tablas de Inventario (fuera del modelo EF — migración manual) ─────────────
using (var scope3 = app.Services.CreateScope())
{
    var cfg3 = scope3.ServiceProvider.GetRequiredService<IConfiguration>();
    await using var conn3 = new Microsoft.Data.SqlClient.SqlConnection(
        cfg3.GetConnectionString("PandoraDb"));
    await conn3.OpenAsync();
    await using var cmd3 = conn3.CreateCommand();
    cmd3.CommandText = """
        IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'InventoryTypes' AND schema_id = SCHEMA_ID('dbo'))
        BEGIN
            CREATE TABLE dbo.InventoryTypes (
                Id          UNIQUEIDENTIFIER NOT NULL DEFAULT NEWID() PRIMARY KEY,
                Name        NVARCHAR(200)    NOT NULL,
                Description NVARCHAR(500)    NULL,
                Department  NVARCHAR(200)    NULL,
                IsActive    BIT              NOT NULL DEFAULT 1,
                CreatedAt   DATETIME2        NOT NULL DEFAULT GETUTCDATE(),
                UpdatedAt   DATETIME2        NULL
            );
        END

        IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'InventoryItems' AND schema_id = SCHEMA_ID('dbo'))
        BEGIN
            CREATE TABLE dbo.InventoryItems (
                Id                 UNIQUEIDENTIFIER NOT NULL DEFAULT NEWID() PRIMARY KEY,
                InventoryNumber    NVARCHAR(50)     NOT NULL,
                Name               NVARCHAR(200)    NOT NULL,
                Brand              NVARCHAR(200)    NULL,
                Model              NVARCHAR(200)    NULL,
                SerialNumber       NVARCHAR(200)    NULL,
                Status             NVARCHAR(50)     NULL DEFAULT 'Activo',
                Department         NVARCHAR(200)    NULL,
                AssignedTo         NVARCHAR(200)    NULL,
                AssignedEmployeeId UNIQUEIDENTIFIER NULL,
                InventoryTypeId    UNIQUEIDENTIFIER NOT NULL,
                IsPhone            BIT              NOT NULL DEFAULT 0,
                IsActive           BIT              NOT NULL DEFAULT 1,
                PurchaseDate       DATETIME2        NULL,
                PurchasePrice      DECIMAL(18,2)    NULL,
                Accessories        NVARCHAR(MAX)    NULL,
                DecommissionDate   DATETIME2        NULL,
                DecommissionReason NVARCHAR(MAX)    NULL,
                CreatedAt          DATETIME2        NOT NULL DEFAULT GETUTCDATE(),
                UpdatedAt          DATETIME2        NULL,
                CONSTRAINT FK_InventoryItems_Types FOREIGN KEY (InventoryTypeId)
                    REFERENCES dbo.InventoryTypes(Id)
            );
        END

        IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'EquipmentTransfers' AND schema_id = SCHEMA_ID('dbo'))
        BEGIN
            CREATE TABLE dbo.EquipmentTransfers (
                Id              UNIQUEIDENTIFIER NOT NULL DEFAULT NEWID() PRIMARY KEY,
                InventoryItemId UNIQUEIDENTIFIER NOT NULL,
                FromDepartment  NVARCHAR(200)    NULL,
                FromPerson      NVARCHAR(200)    NULL,
                ToDepartment    NVARCHAR(200)    NULL,
                ToPerson        NVARCHAR(200)    NULL,
                TransferDate    DATETIME2        NOT NULL DEFAULT GETUTCDATE(),
                Notes           NVARCHAR(MAX)    NULL,
                CreatedBy       NVARCHAR(200)    NULL,
                CreatedAt       DATETIME2        NOT NULL DEFAULT GETUTCDATE(),
                CONSTRAINT FK_EquipmentTransfers_Items FOREIGN KEY (InventoryItemId)
                    REFERENCES dbo.InventoryItems(Id) ON DELETE CASCADE
            );
        END
        """;
    await cmd3.ExecuteNonQueryAsync();
}

await app.RunAsync();

// ── Helpers locales ──────────────────────────────────────────────────────────
static string ResolveLocalDbPipe(string connStr)
{
    if (!connStr.Contains("(localdb)", StringComparison.OrdinalIgnoreCase))
        return connStr;
    try
    {
        var match = Regex.Match(connStr, @"\(localdb\)\\([^;]+)", RegexOptions.IgnoreCase);
        string instance = match.Success ? match.Groups[1].Value.Trim() : "MSSQLLocalDB";
        string info     = RunSqlLocalDb($"info \"{instance}\"");

        if (info.Contains("Detenido") || info.Contains("Stopped") || !info.Contains("pipe", StringComparison.OrdinalIgnoreCase))
        {
            Console.WriteLine($"[LocalDB] Instancia '{instance}' detenida — iniciando...");
            RunSqlLocalDb($"start \"{instance}\"");
            Thread.Sleep(2000);
            info = RunSqlLocalDb($"info \"{instance}\"");
        }

        foreach (string line in info.Split('\n'))
        {
            string t = line.Trim();
            int idx = t.IndexOf(@"\\.\pipe\", StringComparison.OrdinalIgnoreCase);
            if (idx < 0) idx = t.IndexOf("np:", StringComparison.OrdinalIgnoreCase);
            if (idx < 0) continue;

            string pipe = t[idx..].Trim();
            if (pipe.StartsWith("np:", StringComparison.OrdinalIgnoreCase)) pipe = pipe[3..];

            string placeholder = $@"(localdb)\{instance}";
            int pos = connStr.IndexOf(placeholder, StringComparison.OrdinalIgnoreCase);
            string result = pos >= 0
                ? connStr[..pos] + pipe + connStr[(pos + placeholder.Length)..]
                : connStr;
            Console.WriteLine($"[LocalDB] Pipe resuelto: {pipe}");
            return result;
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"[LocalDB] Advertencia al resolver pipe: {ex.Message}");
    }
    return connStr;
}

static string RunSqlLocalDb(string arguments)
{
    var psi = new ProcessStartInfo("sqllocaldb", arguments)
    {
        RedirectStandardOutput = true,
        RedirectStandardError  = true,
        UseShellExecute        = false,
        CreateNoWindow         = true
    };
    using var process = Process.Start(psi)!;
    string output = process.StandardOutput.ReadToEnd();
    process.WaitForExit(8_000);
    return output;
}
