using Microsoft.AspNetCore.SignalR;

namespace Pandora.API.Hubs;

public class NotificationsHub : Hub
{
    /// <summary>Añade la conexión actual al grupo personal del usuario.</summary>
    public Task JoinUser(string username)
        => Groups.AddToGroupAsync(Context.ConnectionId, $"user-{username}");

    /// <summary>Añade la conexión al grupo de difusión global.</summary>
    public Task JoinBroadcast()
        => Groups.AddToGroupAsync(Context.ConnectionId, "broadcast");
}
