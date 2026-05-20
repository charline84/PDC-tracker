import https from 'https';

const get = (url) => {
  https.get(url, (res) => {
    console.log(url, res.statusCode, res.headers.location);
  });
}
get('https://www.data.gouv.fr/api/1/datasets/r/9db13a09-72a9-4871-b430-13872b4890b3');
get('https://www.data.gouv.fr/api/1/datasets/r/8f73cf2d-7bc4-4b5a-b912-718d6991f0a0');
get('https://www.data.gouv.fr/api/1/datasets/r/65a9e264-7a20-46a9-9d98-66becb817bc3');
