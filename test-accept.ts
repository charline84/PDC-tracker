import https from 'https';

const get = (url) => {
  https.get(url, { headers: { 'Accept': 'application/json' } }, (res) => {
    let data = '';
    res.on('data', chunk => data += chunk);
    res.on('end', () => console.log(url, res.statusCode, data.slice(0, 100)));
  });
}
get('https://explore.data.gouv.fr/api/resources/9db13a09-72a9-4871-b430-13872b4890b3/data/json/');
get('https://www.data.gouv.fr/api/resources/9db13a09-72a9-4871-b430-13872b4890b3/data/json/');
