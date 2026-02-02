const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const axios = require('axios');
const fs = require('fs');
const archiver = require('archiver');
const cors = require('cors');
const path = require('path');

const app = express();
const PORT = 5000;

app.use(cors());
app.use(express.static('output'));

const upload = multer({ dest: 'uploads/' });

const ELEVEN_API_KEY = 'sk_ac1f4888ff23b3f1680f0ee1179ab19129a7711aedce3068';
const ELEVEN_VOICE_ID = '5j0HGNGaZNPYPx2FKoZ7';

// Convert text to speech and save audio file
const generateAudio = async (text, filename) => {
  if (!text || typeof text !== 'string' || text.trim() === '') {
    console.log(`Skipping empty or invalid text for ${filename}`);
    return;
  }

  try {
    const response = await axios({
      method: 'post',
      url: `https://api.elevenlabs.io/v1/text-to-speech/${ELEVEN_VOICE_ID}`,
      headers: {
        'xi-api-key': ELEVEN_API_KEY,
        'Content-Type': 'application/json',
      },
      data: {
        text: text.trim(),
        model_id: 'eleven_v3',
        voice_settings: {
          stability: 0.5,
          similarity_boost: 0.75,
        },
      },
      responseType: 'arraybuffer',
    });

    fs.writeFileSync(`audios/${filename}.mp3`, response.data);
  } catch (err) {
    console.error(`Failed to generate audio for "${text}":`, err.response?.data || err.message);
  }
};

// API: Upload Excel and process
app.post('/upload', upload.single('file'), async (req, res) => {
  const workbook = xlsx.readFile(req.file.path);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = xlsx.utils.sheet_to_json(sheet);

  if (!fs.existsSync('audios')) fs.mkdirSync('audios');

  for (let i = 0; i < rows.length; i++) {
    const text = rows[i].text || rows[i].Text || Object.values(rows[i])[0];
    await generateAudio(text, `audio_${i + 1}`);
  }

  // Zip all files
  if (!fs.existsSync('output')) fs.mkdirSync('output');
  const output = fs.createWriteStream('output/audios.zip');
  const archive = archiver('zip');
  archive.pipe(output);
  archive.directory('audios/', false);
  archive.finalize();

  output.on('close', () => {
    res.json({ downloadUrl: 'http://localhost:5000/audios.zip' });
  });
});

app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));