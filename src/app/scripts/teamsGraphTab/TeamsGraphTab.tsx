import * as React from "react";
import { Provider, Flex, Text, Button, Header, Image, Alert } from "@fluentui/react-northstar";
import TeamsBaseComponent, { ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import jwt_decode from "jwt-decode";
import Axios from "axios";
/**
 * State for the teamsGraphTabTab React component
 */
export interface ITeamsGraphTabState extends ITeamsBaseComponentState {
    entityId?: string;
    name?: string;
    error?: string;
    image?: any;
    requireConsent: boolean;
    settings: any;
}

/**
 * Properties for the teamsGraphTabTab React component
 */
export interface ITeamsGraphTabProps {

}

/**
 * Implementation of the teams graph Tab content page
 */
export class TeamsGraphTab extends TeamsBaseComponent<ITeamsGraphTabProps, ITeamsGraphTabState> {

    public async componentWillMount() {
        this.updateTheme(this.getQueryVariable("theme"));


        microsoftTeams.initialize(() => {
            microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);
            microsoftTeams.getContext((context) => {
                this.setState({
                    entityId: context.entityId,
                    requireConsent: false
                });
                this.updateTheme(context.theme);
                microsoftTeams.authentication.getAuthToken({
                    successCallback: (token: string) => {
                        const decoded: { [key: string]: any; } = jwt_decode(token) as { [key: string]: any; };
                        this.setState({ name: decoded!.name });
                        microsoftTeams.appInitialization.notifySuccess();


                        Axios.get(`https://${process.env.HOSTNAME}/api/photo`, {
                            responseType: "blob",
                            headers: {
                                Authorization: `Bearer ${token}`
                            }
                        }).then(result => {
                            // tslint:disable-next-line: no-console
                            const r = new FileReader();
                            r.readAsDataURL(result.data);
                            r.onloadend = () => {
                                if (r.error) {
                                    alert(r.error);
                                } else {
                                    this.setState({ image: r.result });
                                }
                            };
                        }).catch(err => {
                            if (err.message === "Request failed with status code 400") {
                                this.setState({ requireConsent: true });
                            } else {
                                this.setState({ error: err.message });
                            }
                        });


                        Axios.get(`https://${process.env.HOSTNAME}/api/settings/` + context.groupId, {
                            headers: {
                                Authorization: `Bearer ${token}`
                            }
                        }).then(result => {
                            this.setState({ settings: result.data });
                        }).catch(err => {
                            if (err.message === "Request failed with status code 400") {
                                this.setState({ requireConsent: true });
                            } else {
                                this.setState({ error: err.message });
                            }
                        });



                    },
                    failureCallback: (message: string) => {
                        this.setState({ error: message });
                        microsoftTeams.appInitialization.notifyFailure({
                            reason: microsoftTeams.appInitialization.FailedReason.AuthFailed,
                            message
                        });
                    },
                    resources: [process.env.TEAMSGRAPHTAB_APP_URI as string]
                });
            });
        });
    }

    /**
     * The render() method to create the UI of the tab
     */
    public render() {
        return (
            <Provider theme={this.state.theme}>
                <Flex fill={true} column styles={{
                    padding: ".8rem 0 .8rem .5rem"
                }}>
                    <Flex.Item>
                        <Header content="Teams + Graph = ❤" />
                    </Flex.Item>
                    <Flex.Item>
                        <div>
                            {this.state.error &&
                                <Alert content={this.state.error} variables={{ urgent: true }} dismissible />
                            }
                            {this.state.requireConsent &&
                                <Alert warning content="You need to consent to this application" actions={[
                                    {
                                        content: "Consent",
                                        key: "Consent",
                                        primary: true,
                                        onClick: (e) => {
                                            window.open(`https://login.microsoftonline.com/common/adminconsent?client_id=${process.env.TEAMSGRAPHTAB_APP_ID}`);
                                        }
                                    }
                                ]}
                                />}

                            <div>
                                <Image avatar src={this.state.image} styles={{ padding: "5px" }} sizes="largest" />
                                <Text content={`Hello ${this.state.name}`} />
                            </div>

                            <div>
                                {this.state.settings &&
                                    <Alert info content={this.state.settings.funSettings.allowGiphy ? "Giphys are allowed" : "Giphys are NOT allowed"} />
                                }
                            </div>
                        </div>
                    </Flex.Item>
                    <Flex.Item styles={{
                        padding: ".8rem 0 .8rem .5rem"
                    }}>
                        <Text size="smaller" content="(C) Copyright Wictor Wilen" />
                    </Flex.Item>
                </Flex>
            </Provider>
        );
    }
}
