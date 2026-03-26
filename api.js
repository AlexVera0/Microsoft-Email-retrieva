
const express = require('express');
const axios = require('axios');
const cors = require('cors');
const bodyParser = require('body-parser');

const path = require('path');

const app = express();
const port = 3001;

app.use(cors());
app.use(bodyParser.json());

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

async function refreshAccessToken(clientId, refreshToken) {
    const tokenUrl = 'https://login.microsoftonline.com/common/oauth2/v2.0/token';
    const params = new URLSearchParams();
    params.append('client_id', clientId);
    params.append('grant_type', 'refresh_token');
    params.append('refresh_token', refreshToken);
    params.append('scope', 'https://graph.microsoft.com/.default');

    try {
        const response = await axios.post(tokenUrl, params.toString(), {
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
        });
        return response.data.access_token;
    } catch (error) {
        const errorDetail = error.response ? JSON.stringify(error.response.data) : error.message;
        console.error('令牌刷新失败:', errorDetail);
        throw new Error(`授权详细错误: ${errorDetail}`);
    }
}

async function fetchMessages(accessToken, folder = 'inbox') {
    const folderPath = folder === 'junkemail' ? 'junkemail' : 'inbox';
    const apiUrl = `https://graph.microsoft.com/v1.0/me/mailFolders/${folderPath}/messages?$select=id,subject,from,receivedDateTime,body,isRead&$top=50`;

    try {
        const response = await axios.get(apiUrl, {
            headers: { 'Authorization': `Bearer ${accessToken}` }
        });
        return response.data.value;
    } catch (error) {
        console.error('邮件获取失败:', error.response ? error.response.data : error.message);
        throw new Error('获取邮件列表失败，请重试');
    }
}

app.post('/api/ms-mail', async (req, res) => {
    const { clientId, refreshToken, folder } = req.body;

    if (!clientId || !refreshToken) {
        return res.status(400).json({ error: '缺少必要的邮箱信息 (ClientID 或 RefreshToken)' });
    }

    try {
        const accessToken = await refreshAccessToken(clientId, refreshToken);

        const messages = await fetchMessages(accessToken, folder);

        res.json({ success: true, data: messages });
    } catch (error) {
        res.status(500).json({ success: false, error: error.message });
    }
});

app.delete('/api/ms-mail/:id', async (req, res) => {
    const { clientId, refreshToken } = req.body;
    const { id } = req.params;

    if (!clientId || !refreshToken || !id) {
        return res.status(400).json({ error: '缺少必要的邮箱信息或邮件 ID' });
    }

    try {
        const accessToken = await refreshAccessToken(clientId, refreshToken);
        const deleteUrl = `https://graph.microsoft.com/v1.0/me/messages/${id}`;

        await axios.delete(deleteUrl, {
            headers: { 'Authorization': `Bearer ${accessToken}` }
        });

        res.json({ success: true, message: '邮件已移动到已删除邮件或已彻底删除' });
    } catch (error) {
        const errorDetail = error.response ? JSON.stringify(error.response.data) : error.message;
        res.status(500).json({ success: false, error: `删除失败: ${errorDetail}` });
    }
});

app.patch('/api/ms-mail/:id/read', async (req, res) => {
    const { clientId, refreshToken, isRead } = req.body;
    const { id } = req.params;

    if (!clientId || !refreshToken || !id) {
        return res.status(400).json({ error: '缺少必要的邮箱信息或邮件 ID' });
    }

    try {
        const accessToken = await refreshAccessToken(clientId, refreshToken);
        const updateUrl = `https://graph.microsoft.com/v1.0/me/messages/${id}`;

        await axios.patch(updateUrl, { isRead }, {
            headers: { 
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            }
        });

        res.json({ success: true, message: '邮件阅读状态已更新' });
    } catch (error) {
        const errorDetail = error.response ? JSON.stringify(error.response.data) : error.message;
        res.status(500).json({ success: false, error: `状态更新失败: ${errorDetail}` });
    }
});

app.listen(port, '0.0.0.0', () => {
    console.log(`极简后端运行在 http://0.0.0.0:${port}`);
});