const express = require('express');
const cors = require('cors');
const { routes } = require('./routes');

const PORT = 8011;
const message = `Express Server is running on PORT:${PORT}`
const app = express();

app.use(cors());
app.use(express.urlencoded({ extended: false }));
app.use(express.json());
app.use('/', routes);

app.get('/', (req, res) => res.send(message));

app.all('*', (req, res) => res.send('Access denied'));

app.listen(PORT, () => { console.log(message) });
