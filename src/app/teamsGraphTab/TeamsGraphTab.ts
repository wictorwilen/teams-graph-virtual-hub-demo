import { PreventIframe } from "express-msteams-host";

/**
 * Used as place holder for the decorators
 */
@PreventIframe("/teamsGraphTab/index.html")
@PreventIframe("/teamsGraphTab/config.html")
@PreventIframe("/teamsGraphTab/remove.html")
export class TeamsGraphTab {
}
