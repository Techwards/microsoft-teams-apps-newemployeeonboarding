// <copyright file="RemoveNewHireBackgroundService.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.NewHireOnboarding.BackgroundService
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text.RegularExpressions;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Hosting;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.NewHireOnboarding.Helpers;
    using Microsoft.Teams.Apps.NewHireOnboarding.Models;
    using Microsoft.Teams.Apps.NewHireOnboarding.Models.Configuration;
    using Microsoft.Teams.Apps.NewHireOnboarding.Models.EntityModels;
    using Microsoft.Teams.Apps.NewHireOnboarding.Providers;

    /// <summary>
    /// A Background service which checks the New Hire Duration after every day. If the duration expires the New Hire record gets
    /// deleted and the app gets deleted...
    /// </summary>
    public class RemoveNewHireBackgroundService : BackgroundService
    {
        private readonly ILogger<RemoveNewHireBackgroundService> logger;

        /// <summary>
        /// A set of key/value application configuration properties for bot settings.
        /// </summary>
        private readonly IOptions<BotOptions> botOptions;

        /// <summary>
        /// Instance to work with Microsoft Graph methods.
        /// </summary>
        private readonly IGraphUtilityHelper graphTokenUtility;

        /// <summary>
        /// Gets configuration setting for New Hire Retention Period in Days.
        /// </summary>
        private readonly IOptionsMonitor<RemoveNewHireBackgroundServiceSettings> removeNewHireBackgroundServiceOption;

        /// <summary>
        /// Provider for fetching information about user details from storage.
        /// </summary>
        private readonly IUserStorageProvider userStorageProvider;

        /// <summary>
        /// Helper for team operations with Microsoft Graph API.
        /// </summary>
        private readonly ITeamMembership membersService;

        /// <summary>
        /// Initializes a new instance of the <see cref="RemoveNewHireBackgroundService"/> class.
        /// BackgroundService class that inherits IHostedService and implements the methods related to deleting New Hire record tasks.
        /// </summary>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="backgroundServiceOption">Instance to get NewHireRetentionPeriodInDays value.</param>
        /// <param name="userStorageProvider">Provider for fetching information about user details from storage.</param>
        /// <param name="botOptions">A set of key/value application configuration properties.</param>
        /// <param name="graphTokenUtility">Instance of Microsoft Graph utility helper.</param>
        /// <param name="teamMembershipHelper">Helper for team operations with Microsoft Graph API.</param>
        public RemoveNewHireBackgroundService(
            ILogger<RemoveNewHireBackgroundService> logger,
            IOptionsMonitor<RemoveNewHireBackgroundServiceSettings> backgroundServiceOption,
            IUserStorageProvider userStorageProvider,
            IOptions<BotOptions> botOptions,
            IGraphUtilityHelper graphTokenUtility,
            ITeamMembership teamMembershipHelper)
        {
            this.logger = logger ?? throw new ArgumentNullException(nameof(logger));
            this.removeNewHireBackgroundServiceOption = backgroundServiceOption ?? throw new ArgumentNullException(nameof(backgroundServiceOption));
            this.userStorageProvider = userStorageProvider ?? throw new ArgumentNullException(nameof(userStorageProvider));
            this.graphTokenUtility = graphTokenUtility ?? throw new ArgumentNullException(nameof(graphTokenUtility));
            this.botOptions = botOptions ?? throw new ArgumentNullException(nameof(botOptions));
            this.membersService = teamMembershipHelper ?? throw new ArgumentNullException(nameof(teamMembershipHelper));
        }

        /// <summary>
        /// This method is called when the Microsoft.Extensions.Hosting.IHostedService starts.
        /// The implementation should return a task that represents the lifetime of the long
        /// running operation(s) being performed.
        /// </summary>
        /// <param name="stoppingToken">Triggered when Microsoft.Extensions.Hosting.IHostedService. StopAsync(System.Threading.CancellationToken) is called.</param>
        /// <returns>A System.Threading.Tasks.Task that represents the long running operations.</returns>
        protected async override Task ExecuteAsync(CancellationToken stoppingToken)
        {
            while (!stoppingToken.IsCancellationRequested)
            {
                try
                {
                    var currentDateTime = DateTime.UtcNow;
                    this.logger.LogInformation($"Remove New Hire Background Service starts running at: {currentDateTime}.");

                    var response = await this.graphTokenUtility.ObtainApplicationTokenAsync(
                        this.botOptions.Value.TenantId, this.botOptions.Value.MicrosoftAppId, this.botOptions.Value.MicrosoftAppPassword);
                    if (response == null)
                    {
                        this.logger.LogInformation($"Failed to acquire application token for application Id: {this.botOptions.Value.MicrosoftAppId}.");
                        return;
                    }

                    this.logger.LogInformation("Checking Employee Experience time period...");
                    List<UserEntity> employeesToBeRemoved = await this.BrowseNewHireDurationAsync();

                    if (employeesToBeRemoved != null)
                    {
                        await this.RemoveAppFromUserScopeAsync(response.AccessToken, employeesToBeRemoved);
                    }
                }
#pragma warning disable CA1031 // Do not catch general exception types
                catch (Exception ex)
#pragma warning restore CA1031 // Do not catch general exception types
                {
                    this.logger.LogError(ex, $"Error while removing Employee. Exception : {ex.Message}");
                }
                finally
                {
                    this.logger.LogInformation("Resume after 5 seconds...");
                    await Task.Delay(TimeSpan.FromSeconds(5), stoppingToken);
                }
            }
        }

        /// <summary>
        /// This method removes the New Hire Data from the database
        /// </summary>
        /// <returns>None.</returns>
        private async Task<List<UserEntity>> BrowseNewHireDurationAsync()
        {
            var currentTime = DateTime.UtcNow;

            var newHires = await this.userStorageProvider.GetAllUsersAsync(UserRole.NewHire);
            if (newHires == null || !newHires.Any())
            {
                this.logger.LogError("New hires not available.");
                return null;
            }

            var employeesToBeRemoved = newHires.Where(
                employee => (currentTime - employee.BotInstalledOn)?.Days > this.removeNewHireBackgroundServiceOption.CurrentValue.NewHireRetentionPeriodInDays).ToList();
            if (!employeesToBeRemoved.Any())
            {
                this.logger.LogInformation("No New Hires completed their retention period.");
                return null;
            }

            return employeesToBeRemoved;
        }

        /// <summary>
        /// Remove New Hire Data from Azure Storage Table and eventually App from user scope.
        /// </summary>
        /// <param name="graphApiAccessToken">Application type Access Token for Graph API.</param>
        /// <param name="employeesToBeRemoved">Contains New Hire Data (that to be removed) from UserEntity Storage Table.</param>
        /// <returns>None.</returns>
        private async Task RemoveAppFromUserScopeAsync(string graphApiAccessToken, List<UserEntity> employeesToBeRemoved)
        {
            // Todo: Perhaps we need to remove all records that corresponds to the user aad Id.
            // Currently, the data from User Configuration Table gets deleted.
            await this.userStorageProvider.DeleteUserRecordsBatchAsync(employeesToBeRemoved);

            List<string> employeeAadIdList = new List<string>();
            foreach (var employeeEntity in employeesToBeRemoved)
            {
                employeeAadIdList.Add(employeeEntity.AadObjectId);
            }

            // Perhaps we dont need below commented out method anymore... If this is the case we need to remove the "TeamsLink" parameter
            // from the BotOptions Configuration Model, from the "azuredeploy.json" file (from Deployment directory), from "ServicesExtension.cs"
            // file and from "appsettings.json" file.
            // string teamsId = this.GetPatternMatchedValue(this.botOptions.Value.TeamsLink, @"groupId=([-\w]+?)&");
            if (employeeAadIdList.Any())
            {
                foreach (var employeeAadId in employeeAadIdList)
                {
                    string installedAppId = await this.membersService.GetInstalledAppIdAsync(graphApiAccessToken, employeeAadId);
                    await this.membersService.RemoveAppFromUserScopeAsync(graphApiAccessToken, employeeAadId, installedAppId);
                }
            }
        }

        /// <summary>
        /// Get the required value from the targetString using Regex pattern.
        /// </summary>
        /// <param name="targetString">String on which the Regex pattern has to be applied.</param>
        /// <param name="pattern">Regex pattern.</param>
        /// <returns>Return the required string.</returns>
        private string GetPatternMatchedValue(string targetString, string pattern)
        {
            Match match = Regex.Match(targetString, pattern);
            return match.Groups[1].Value;
        }
    }
}
