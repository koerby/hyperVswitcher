using System.Collections.Concurrent;

namespace HyperTool.Services;

public static class UsbGuestConnectionRegistry
{
    private sealed class GuestConnectionEntry
    {
        public string GuestComputerName { get; init; } = string.Empty;
        public DateTimeOffset LastSeenUtc { get; init; }
    }

    private static readonly ConcurrentDictionary<string, GuestConnectionEntry> ConnectedGuestsByBusId = new(StringComparer.OrdinalIgnoreCase);

    public static void UpdateFromDiagnosticsAck(HyperVSocketDiagnosticsAck ack)
    {
        if (ack is null)
        {
            return;
        }

        var busId = (ack.BusId ?? string.Empty).Trim();
        var eventType = (ack.EventType ?? string.Empty).Trim();
        var guestComputerName = (ack.GuestComputerName ?? string.Empty).Trim();

        if (string.IsNullOrWhiteSpace(busId))
        {
            return;
        }

        if (string.Equals(eventType, "usb-disconnected", StringComparison.OrdinalIgnoreCase))
        {
            ConnectedGuestsByBusId.TryRemove(busId, out _);
            return;
        }

        if ((string.Equals(eventType, "usb-connected", StringComparison.OrdinalIgnoreCase)
             || string.Equals(eventType, "usb-heartbeat", StringComparison.OrdinalIgnoreCase))
            && !string.IsNullOrWhiteSpace(guestComputerName))
        {
            ConnectedGuestsByBusId[busId] = new GuestConnectionEntry
            {
                GuestComputerName = guestComputerName,
                LastSeenUtc = DateTimeOffset.UtcNow
            };
        }
    }

    public static bool TryGetGuestComputerName(string? busId, out string guestComputerName)
    {
        guestComputerName = string.Empty;

        if (string.IsNullOrWhiteSpace(busId))
        {
            return false;
        }

        if (!ConnectedGuestsByBusId.TryGetValue(busId.Trim(), out var entry))
        {
            return false;
        }

        guestComputerName = entry.GuestComputerName;
        return true;
    }

    public static bool TryGetFreshGuestComputerName(string? busId, TimeSpan maxAge, out string guestComputerName)
    {
        guestComputerName = string.Empty;

        if (string.IsNullOrWhiteSpace(busId))
        {
            return false;
        }

        if (!ConnectedGuestsByBusId.TryGetValue(busId.Trim(), out var entry))
        {
            return false;
        }

        if ((DateTimeOffset.UtcNow - entry.LastSeenUtc) > maxAge)
        {
            return false;
        }

        guestComputerName = entry.GuestComputerName;
        return true;
    }
}
