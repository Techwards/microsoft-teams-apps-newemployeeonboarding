// <copyright file="RemoveNewHireBackgroundServiceSettings.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.NewHireOnboarding.Models.Configuration
{
    /// <summary>
    /// This class used to set value of New Hire retention period which is then used by Remove New Hire background service.
    /// </summary>
    public class RemoveNewHireBackgroundServiceSettings
    {
        /// <summary>
        /// Gets or sets New Hire retention period.
        /// </summary>
        public int NewHireRetentionPeriodInDays { get; set; }
    }
}
