using Microsoft.AspNetCore.SignalR;

namespace Pandora.API.Hubs;

public class ProgressHub : Hub
{
    public Task JoinCampaign(string campaignId)
        => Groups.AddToGroupAsync(Context.ConnectionId, campaignId);

    public Task LeaveCampaign(string campaignId)
        => Groups.RemoveFromGroupAsync(Context.ConnectionId, campaignId);
}
