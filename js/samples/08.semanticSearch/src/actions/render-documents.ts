import debug from 'debug';
import { TurnState } from '@microsoft/teams-ai';
import { TurnContext, CardFactory, Attachment } from 'botbuilder';

import { SemanticSearchClient } from '../client';

interface RenderDocumentsActionParams {
    readonly documents: {
        readonly id: string;
        readonly siteId: string;
        readonly name: string;
        readonly summary?: string;
    }[];
}

export function renderDocuments() {
    const log = debug('m365:actions:RenderDocuments');

    return async (context: TurnContext, state: TurnState, params: RenderDocumentsActionParams): Promise<string> => {
        const token = state.temp.authTokens['graph'];

        if (!token) {
            return 'no auth token found';
        }

        log(params);

        const client = new SemanticSearchClient(token);
        const attachments: Attachment[] = [];

        for (const document of params.documents) {
            const item = (await client.getDriveItem(document.siteId, document.id)).data;
            const thumbnails = (await client.getDriveItemThumbnails(document.siteId, document.id)).data;
            const thumbnail = thumbnails.value[0].small || thumbnails.value[0].medium || thumbnails.value[0].large

            attachments.push(CardFactory.adaptiveCard({
                $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
                type: 'AdaptiveCard',
                version: '1.4',
                body: [
                    {
                        type: 'ColumnSet',
                        horizontalAlignment: 'left',
                        columns: [
                            {
                                type: 'Column',
                                items: [
                                    {
                                        type: 'Image',
                                        url: thumbnail.url,
                                        height: `${thumbnail.height}px`,
                                        width: `${thumbnail.width}px`
                                    }
                                ]
                            },
                            {
                                type: 'Column',
                                items: [
                                    {
                                        type: 'TextBlock',
                                        text: item.name,
                                        size: 'large',
                                        weight: 'bolder',
                                        horizontalAlignment: 'left'
                                    },
                                    {
                                        type: 'TextBlock',
                                        text: document.summary,
                                        wrap: true,
                                        maxLines: 3,
                                        isSubtle: true,
                                        horizontalAlignment: 'left'
                                    }
                                ]
                            }
                        ]
                    }
                ],
                actions: [
                    {
                        type: 'Action.OpenUrl',
                        url: item.webUrl,
                        title: "Open"
                    },
                    {
                        type: 'Action.OpenUrl',
                        url: item['@microsoft.graph.downloadUrl'],
                        title: "Download"
                    }
                ]
            }));
        }

        await context.sendActivity({ attachments });
        return '';
    };
}
