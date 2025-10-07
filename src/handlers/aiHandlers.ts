export async function handleAITableGeneration(msg: any) {
    try {
        const { prompt, apiKey, rows, cols } = msg as { prompt: string, apiKey: string, rows: number, cols: number };
        const proxyUrl = 'https://table-generator-server.vercel.app';
        const endpoint = 'https://us-south.ml.cloud.ibm.com';

        const tokenRes = await fetch(`${proxyUrl}/token`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ apiKey })
        });
        if (!tokenRes.ok) {
            const t = await tokenRes.text();
            throw new Error(`proxy token error: ${tokenRes.status} ${t}`);
        }
        const { access_token } = await tokenRes.json();

        const tableRes = await fetch(`${proxyUrl}/generateTable`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ endpoint, accessToken: access_token, prompt, rows, cols })
        });
        if (!tableRes.ok) {
            const t = await tableRes.text();
            throw new Error(`proxy generateTable error: ${tableRes.status} ${t}`);
        }
        const tableJson = await tableRes.json();
        const headers: string[] = Array.isArray(tableJson.headers) ? tableJson.headers.map(String).slice(0, cols) : Array(cols).fill('').map((_, i) => `Column ${i+1}`);
        const rowsData: string[][] = Array.isArray(tableJson.rows) ? tableJson.rows.map((r: any) => Array.isArray(r) ? r.map(String) : []) : [];
        figma.ui.postMessage({ type: 'ai-table-response', headers, rows: rowsData });
    } catch (e) {
        console.error('generate-table-with-ai error', e);
        figma.ui.postMessage({ type: 'ai-table-response', headers: [], rows: [] });
        figma.notify('Failed to generate table with AI');
    }
}

export async function handleAIContentGeneration(msg: any) {
    console.log('🤖 AI content generation handler called!', msg);
    try {
        const { prompt, cellKey } = msg;
        console.log('Generating AI content for:', cellKey, 'with prompt:', prompt);
        
        // Immediate response for testing
        const randomResponse = 'AI Generated: ' + prompt;
        figma.ui.postMessage({
            type: 'ai-content-generated',
            cellKey,
            content: randomResponse
        });
        
    } catch (error) {
        console.error('Error generating AI content:', error);
        figma.ui.postMessage({
            type: 'ai-content-error',
            cellKey: msg.cellKey,
            error: 'Failed to generate AI content'
        });
    }
}

export async function handleWatsonxDataGeneration(msg: any) {
    try {
        const { prompt, endpoint, apiKey, accessToken: uiAccessToken, useAccessToken, useProxy, proxyUrl, count, remember } = msg as { prompt: string, endpoint: string, apiKey: string, accessToken?: string, useAccessToken?: boolean, useProxy?: boolean, proxyUrl?: string, count: number, remember?: boolean };
        let endpointToUse = endpoint;
        let apiKeyToUse = apiKey;
        
        // Fallback to stored values if missing
        if (!endpointToUse) {
            const storedEndpoint = await figma.clientStorage.getAsync('watsonx.endpoint');
            if (storedEndpoint) endpointToUse = String(storedEndpoint);
        }
        if (!apiKeyToUse) {
            const storedKey = await figma.clientStorage.getAsync('watsonx.apiKey');
            if (storedKey) apiKeyToUse = String(storedKey);
        }
        if (remember) {
            try {
                if (endpoint) await figma.clientStorage.setAsync('watsonx.endpoint', endpoint);
                if (apiKey) await figma.clientStorage.setAsync('watsonx.apiKey', apiKey);
            } catch (e) {
                console.warn('Failed to persist watsonx settings', e);
            }
        }
        if (!endpointToUse || (!useAccessToken && !apiKeyToUse && !uiAccessToken)) {
            figma.ui.postMessage({ type: 'watsonx-data-response', data: [] });
            figma.notify('Missing watsonx endpoint or credentials');
            return;
        }
        
        let accessToken = (useAccessToken && uiAccessToken) ? uiAccessToken : '';
        if (useProxy && proxyUrl) {
            console.log('Using proxy server:', proxyUrl);
            if (!accessToken) {
                console.log('Getting token via proxy...');
                const tokenRes = await fetch(`${proxyUrl.replace(/\/$/, '')}/token`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ apiKey: apiKeyToUse })
                });
                if (!tokenRes.ok) throw new Error('Proxy token fetch failed');
                const tokenJson = await tokenRes.json();
                accessToken = tokenJson.access_token as string;
                console.log('Got token via proxy');
            }
            console.log('Generating text via proxy...');
            const genRes = await fetch(`${proxyUrl.replace(/\/$/, '')}/generate`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ endpoint: endpointToUse, accessToken, prompt, count })
            });
            if (!genRes.ok) {
                const t = await genRes.text();
                throw new Error(`proxy watsonx error: ${genRes.status} ${t}`);
            }
            const genJson = await genRes.json();
            console.log('Proxy response:', genJson);
            const lines = genJson.data || [];
            figma.ui.postMessage({ type: 'watsonx-data-response', data: lines });
        } else {
            console.log('Using direct API calls (may fail due to CORS)');
            if (!accessToken) {
                const tokenRes = await fetch('https://iam.cloud.ibm.com/identity/token', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Accept': 'application/json' },
                    body: `grant_type=urn:ibm:params:oauth:grant-type:apikey&apikey=${encodeURIComponent(apiKeyToUse)}`
                });
                if (!tokenRes.ok) throw new Error('Failed to fetch IAM token');
                const tokenJson = await tokenRes.json();
                accessToken = tokenJson.access_token as string;
            }
            const genUrl = `${endpointToUse.replace(/\/$/, '')}/ml/v1/text/chat?version=2023-05-29`;
            const wxBody = {
                input: `${prompt}\nReturn ${count} short values as a plain list, one per line, no numbering.`,
                parameters: { decoding_method: 'greedy', max_new_tokens: 16, stop_sequences: ['\n\n'] }
            } as any;
            const genRes = await fetch(genUrl, { method: 'POST', headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${accessToken}` }, body: JSON.stringify(wxBody) });
            if (!genRes.ok) {
                const t = await genRes.text();
                throw new Error(`watsonx error: ${genRes.status} ${t}`);
            }
            const genJson = await genRes.json();
            let text = '';
            if (Array.isArray(genJson.results) && genJson.results[0]?.generated_text) text = genJson.results[0].generated_text as string;
            else if (genJson.generated_text) text = genJson.generated_text as string;
            else text = String(genJson.output || '');
            const lines = text.split('\n').map((s: string) => s.replace(/^[-*\d\.\)\s]+/, '').trim()).filter((s: string) => s.length > 0).slice(0, count);
            figma.ui.postMessage({ type: 'watsonx-data-response', data: lines });
        }
    } catch (err: any) {
        figma.ui.postMessage({ type: 'watsonx-data-response', data: [] });
        figma.notify('watsonx.ai request failed');
        console.error('watsonx.ai error', err);
    }
}