using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NwScheduleFormatter.Models;

public class JwWebsiteMeeting
{
    public Song InitialSong { get; set; } = new();
    public Song MiddleSong { get; set; } = new();
    public Song FinalSong { get; set; } = new();

    public short Apply1DurationMinutes { get; set; }
    public short Apply2DurationMinutes { get; set; }
    public short Apply3DurationMinutes { get; set; }
    public short? Apply4DurationMinutes { get; set; }

    public short LivingAsChristhians1DurationMinutes { get; set; }
    public short? LivingAsChristhians2DurationMinutes { get; set; }

}
