using System;
using System.Collections.Generic;
using System.Net;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Web;
using Microsoft.Identity.Web.Resource;
using Microsoft.Graph;
using TimeZoneConverter;

namespace GraphTutorial.Controllers
{
    [ApiController]
    [Route("[controller]")]
    [Authorize]
    public class CalendarController : ControllerBase
    {
        private static readonly string[] apiScopes = new[] { "access_as_user" };

        private readonly GraphServiceClient _graphClient;
        private readonly ITokenAcquisition _tokenAcquisition;
        private readonly ILogger<CalendarController> _logger;

        public CalendarController(ITokenAcquisition tokenAcquisition, GraphServiceClient graphClient, ILogger<CalendarController> logger)
        {
            _tokenAcquisition = tokenAcquisition;
            _graphClient = graphClient;
            _logger = logger;
        }

        [HttpGet]
        public async Task<ActionResult<string>> Get()
        {
            // This verifies that the access_as_user scope is
            // present in the bearer token, throws if not
            HttpContext.VerifyUserHasAnyAcceptedScope(apiScopes);

            // To verify that the identity libraries have authenticated
            // based on the token, log the user's name
            _logger.LogInformation($"Authenticated user: {User.GetDisplayName()}");

            try
            {
                // TEMPORARY
                // Get a Graph token via OBO flow
                var token = await _tokenAcquisition
                    .GetAccessTokenForUserAsync(new[]{
                        "User.Read",
                        "MailboxSettings.Read",
                        "Calendars.ReadWrite" });

                // Log the token
                _logger.LogInformation($"Access token for Graph: {token}");
                return Ok("{ \"status\": \"OK\" }");
            }
            catch (MicrosoftIdentityWebChallengeUserException ex)
            {
                _logger.LogError(ex, "Consent required");
                // This exception indicates consent is required.
                // Return a 403 with "consent_required" in the body
                // to signal to the tab it needs to prompt for consent
                return new ContentResult {
                    StatusCode = (int)HttpStatusCode.Forbidden,
                    ContentType = "text/plain",
                    Content = "consent_required"
                };
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error occurred");
                throw;
            }
        }
    }
}