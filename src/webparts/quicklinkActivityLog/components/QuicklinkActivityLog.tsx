import * as React from 'react';
import { useState, useEffect } from 'react';
import { getSP } from '../../../pnpjsConfig';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Stack, Label, MessageBar, MessageBarType } from '@fluentui/react';
import "./QuicklinkActivityLog.css";

interface IQuickLink {
    id: number;
    title: string;
    url: string;
    pageName: string;
    webPartTitle: string;
}

interface IProps {
    context: WebPartContext;
    webPartTitle: string;
}

const QuicklinkActivityLog: React.FC<IProps> = ({ context, webPartTitle }) => {
    const [links, setLinks] = useState<IQuickLink[]>([]);
    const [error, setError] = useState<string>('');

    useEffect(() => {
        if (!webPartTitle) {
            setError("Web Part title is required.");
            return;
        }
        loadLinks();
    }, [webPartTitle]);

    const loadLinks = async () => {
        try {
            const sp = getSP(context);
            const items = await sp.web.lists.getByTitle('QuickLinksData').items.filter(`PageName eq '${context.pageContext.web.title}' and WebPartTitle eq '${webPartTitle}'`)();
            const formattedLinks = items.map((item: any) => ({
                id: item.Id,
                title: item.Title,
                url: item.LinkUrl,
                pageName: item.PageName,
                webPartTitle: item.WebPartTitle
            }));
            setLinks(formattedLinks);
        } catch (err) {
            console.error("Error loading links from SharePoint:", err);
            setError("Failed to load links. Please try again later.");
        }
    };

    const logClickEvent = async (link: IQuickLink) => {
        try {
            const sp = getSP(context);
            const user = await sp.web.currentUser();
            await sp.web.lists.getByTitle('ClickAnalytics').items.add({
                Title: link.title,
                LinkUrl: link.url,
                User: user.Title,
                ClickedTime: new Date().toISOString()
            });
        } catch (err) {
            console.error("Error logging click event:", err);
        }
    };

    return (
        <div className="container-fluid">
            <Label>{webPartTitle || 'Quick Links'}</Label>
            {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}

            <Stack tokens={{ childrenGap: 10 }}>
                {links.length > 0 ? (
                    links.map((link) => (
                        <div key={link.id} className="linkCard">
                            <a 
                                href={link.url}
                                target="_blank"
                                rel="noopener noreferrer"
                                onClick={() => logClickEvent(link)}
                            >
                                {link.title}
                            </a>
                        </div>
                    ))
                ) : (
                    <p>No links available.</p>
                )}
            </Stack>
        </div>
    );
};

export default QuicklinkActivityLog;