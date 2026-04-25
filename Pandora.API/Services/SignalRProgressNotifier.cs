using Microsoft.AspNetCore.SignalR;
using Pandora.API.Hubs;
using Pandora.Application.Interfaces;

namespace Pandora.API.Services;

public class SignalRProgressNotifier(IHubContext<ProgressHub> hub) : IProgressNotifier
{
    public Task NotifyAsync(Guid campaignId, int current, int total)
        => hub.Clients.Group(campaignId.ToString())
               .SendAsync("progress", new { current, total });
}
