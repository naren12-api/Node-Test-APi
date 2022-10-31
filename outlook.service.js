const request = require("request");
const config = require("./config/config");

const cookie = config.cookie;
const xOwacanary = config.owaCanary;

const requiredOptions = {
    method: 'POST',
    url: 'https://outlook.live.com/owa/0/service.svc',
    qs: { action: '', app: 'Mail', n: '115' },
    headers:
    {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:98.0) Gecko/20100101 Firefox/98.0',
        'cache-control': 'no-cache',
        "cookie": cookie,
        "x-owa-canary": xOwacanary,
        "accept": "*/*",
        "accept-language": "en-US,en;q=0.9",
        "action": "",
        "content-type": "application/json; charset=utf-8",
        "Referer": "https://outlook.live.com/",
        "Referrer-Policy": "strict-origin"
    },
}

const doRequest = (options) => {
    const updatedOptions = { ...requiredOptions, headers: {...requiredOptions.headers, ...options.headers}, body: options?.body, qs: {...requiredOptions.qs, ...options.qs} };
    return new Promise(function (resolve, reject) {
        request(updatedOptions, function (error, res, body) {
            if (!error && res.statusCode == 200) {
                resolve(body);
            } else {
                if (error) {
                    reject(error);
                } else {
                    reject(body);
                }
            }
        });
    });
}

module.exports = { doRequest }