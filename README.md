# playlist2excel

Python script that extracts Spotify playlists and exports them as an Excel file, with each playlist as a separate sheet.

## Features

- 🎵 Extracts all playlists from your Spotify account
- 📊 Exports to Excel with each playlist as a separate sheet
- 🔒 Secure credential management using environment variables
- 📝 Detailed track information (name, artist, album, release date, popularity, etc.)

## Prerequisites

- Python 3.7 or higher
- A Spotify account
- Spotify Developer App credentials

## Setup

### 1. Get Spotify API Credentials

1. Go to [Spotify Developer Dashboard](https://developer.spotify.com/dashboard)
2. Log in with your Spotify account
3. Click "Create an App"
4. Fill in the app name and description
5. Copy your **Client ID** and **Client Secret**
6. Click "Edit Settings" and add `http://127.0.0.1:8888/callback` to the Redirect URIs

### 2. Install Dependencies

```bash
pipx install -r requirements.txt
```

### 3. Configure Environment Variables

Create a `.env` file in the project directory (or set environment variables in your system):

```bash
cp .env.example .env
```

Edit `.env` and add your credentials:

```
SPOTIFY_CLIENT_ID=your_client_id_here
SPOTIFY_CLIENT_SECRET=your_client_secret_here
SPOTIFY_REDIRECT_URI=http://127.0.0.1:8888/callback
```

**Note:** The `.env` file is gitignored to prevent accidentally committing your credentials.

## Usage

### Make the script executable (Linux/Mac)

```bash
chmod +x playlist2excel.py
```

### Run the script

Using environment variables from `.env` file:

```bash
# Load environment variables from .env
export $(cat .env | xargs)
./playlist2excel.py
```

Or set them inline:

```bash
SPOTIFY_CLIENT_ID=your_id SPOTIFY_CLIENT_SECRET=your_secret ./playlist2excel.py
```

Or using Python directly:

```bash
python3 playlist2excel.py
```

### First Run

On the first run, a browser window will open asking you to authorize the app. After authorization, you'll be redirected to a URL. Copy the entire URL and paste it back into the terminal when prompted.

### Output

The script creates a file named `spotify_playlists.xlsx` containing:
- Each playlist as a separate sheet
- Track details: name, artist(s), album, release date, duration, popularity, date added, and Spotify URL

## Troubleshooting

**Error: Missing Spotify credentials**
- Ensure your environment variables are set correctly
- Check that `.env` file exists and contains valid credentials

**Error: redirect_uri mismatch**
- Ensure the redirect URI in your Spotify app settings matches the one in your `.env` file

**Browser doesn't open automatically**
- Copy the URL printed in the terminal and paste it into your browser manually

## License

This project is open source and available under the MIT License.