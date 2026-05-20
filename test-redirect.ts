import https from 'https';

const get = (url) => {
  https.get(url, (res) => {
    console.log(url, res.statusCode, res.headers.location);
  });
}
get('https://data.gouv.fr/api/1/resources/9db13a09-72a9-4871-b430-13872b4890b3/data/json/');
