import https from 'https';
https.get('https://www.data.gouv.fr/api/resources/9db13a09-72a9-4871-b430-13872b4890b3/data/json/', (res) => {
  console.log('Status Code:', res.statusCode);
  let data = '';
  res.on('data', chunk => data += chunk);
  res.on('end', () => console.log('Data:', data.slice(0, 200)));
});
