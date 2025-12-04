#!/usr/bin/env python3
"""
Spotify Playlist to Excel Exporter
Extracts playlists from Spotify and creates an Excel file with each playlist as a separate sheet.
"""

import os
import sys
from typing import List, Dict, Any
import spotipy
from spotipy.oauth2 import SpotifyOAuth
import pandas as pd
from openpyxl.utils.exceptions import IllegalCharacterError


def get_spotify_client() -> spotipy.Spotify:
    """
    Initialize and return a Spotify client using credentials from environment variables.
    
    Required environment variables:
    - SPOTIFY_CLIENT_ID: Your Spotify app client ID
    - SPOTIFY_CLIENT_SECRET: Your Spotify app client secret
    - SPOTIFY_REDIRECT_URI: Your Spotify app redirect URI (default: http://localhost:8888/callback)
    
    Returns:
        spotipy.Spotify: Authenticated Spotify client
    """
    client_id = os.getenv('SPOTIFY_CLIENT_ID')
    client_secret = os.getenv('SPOTIFY_CLIENT_SECRET')
    redirect_uri = os.getenv('SPOTIFY_REDIRECT_URI', 'http://localhost:8888/callback')
    
    if not client_id or not client_secret:
        print("Error: Missing Spotify credentials!")
        print("Please set the following environment variables:")
        print("  - SPOTIFY_CLIENT_ID")
        print("  - SPOTIFY_CLIENT_SECRET")
        print("  - SPOTIFY_REDIRECT_URI (optional, default: http://localhost:8888/callback)")
        sys.exit(1)
    
    scope = "playlist-read-private playlist-read-collaborative"
    
    auth_manager = SpotifyOAuth(
        client_id=client_id,
        client_secret=client_secret,
        redirect_uri=redirect_uri,
        scope=scope
    )
    
    return spotipy.Spotify(auth_manager=auth_manager)


def get_user_playlists(sp: spotipy.Spotify) -> List[Dict[str, Any]]:
    """
    Fetch all playlists for the authenticated user.
    
    Args:
        sp: Authenticated Spotify client
        
    Returns:
        List of playlist dictionaries
    """
    playlists = []
    results = sp.current_user_playlists(limit=50)
    
    while results:
        playlists.extend(results['items'])
        if results['next']:
            results = sp.next(results)
        else:
            break
    
    return playlists


def get_playlist_tracks(sp: spotipy.Spotify, playlist_id: str) -> List[Dict[str, Any]]:
    """
    Fetch all tracks from a specific playlist.
    
    Args:
        sp: Authenticated Spotify client
        playlist_id: Spotify playlist ID
        
    Returns:
        List of track information dictionaries
    """
    tracks = []
    results = sp.playlist_tracks(playlist_id, limit=100)
    
    while results:
        for item in results['items']:
            if item['track'] is None:
                continue
                
            track = item['track']
            track_info = {
                'Track Name': track.get('name', 'N/A'),
                'Artist(s)': ', '.join([artist['name'] for artist in track.get('artists', [])]),
                'Album': track.get('album', {}).get('name', 'N/A'),
                'Release Date': track.get('album', {}).get('release_date', 'N/A'),
                'Duration (ms)': track.get('duration_ms', 0),
                'Popularity': track.get('popularity', 0),
                'Added At': item.get('added_at', 'N/A'),
                'Added By': item.get('added_by', {}).get('id', 'N/A'),
                'Spotify URL': track.get('external_urls', {}).get('spotify', 'N/A')
            }
            tracks.append(track_info)
        
        if results['next']:
            results = sp.next(results)
        else:
            break
    
    return tracks


def sanitize_sheet_name(name: str) -> str:
    """
    Sanitize playlist name to be a valid Excel sheet name.
    Excel sheet names can't exceed 31 characters and can't contain: : \\ / ? * [ ]
    
    Args:
        name: Original playlist name
        
    Returns:
        Sanitized sheet name
    """
    # Remove invalid characters
    invalid_chars = [':', '\\', '/', '?', '*', '[', ']']
    for char in invalid_chars:
        name = name.replace(char, '')
    
    # Truncate to 31 characters
    if len(name) > 31:
        name = name[:31]
    
    # If name is empty after sanitization, use a default
    if not name.strip():
        name = "Unnamed Playlist"
    
    return name


def export_to_excel(playlists_data: Dict[str, pd.DataFrame], output_file: str = 'spotify_playlists.xlsx'):
    """
    Export playlists to an Excel file with each playlist as a separate sheet.
    
    Args:
        playlists_data: Dictionary mapping playlist names to DataFrames
        output_file: Output Excel file name
    """
    if not playlists_data:
        print("No playlists to export!")
        return
    
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for sheet_name, df in playlists_data.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print(f"\n✓ Successfully exported {len(playlists_data)} playlists to '{output_file}'")
        
    except IllegalCharacterError as e:
        print(f"Error: Illegal character in data: {e}")
        print("Trying to export with cleaned data...")
        
        # Clean the data and retry
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for sheet_name, df in playlists_data.items():
                # Clean string columns
                for col in df.select_dtypes(include=['object']).columns:
                    df[col] = df[col].apply(lambda x: ''.join(char for char in str(x) if ord(char) >= 32) if pd.notna(x) else x)
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print(f"\n✓ Successfully exported {len(playlists_data)} playlists to '{output_file}' (with cleaned data)")


def main():
    """Main function to orchestrate the playlist export process."""
    print("Spotify Playlist to Excel Exporter")
    print("=" * 50)
    
    # Initialize Spotify client
    print("\n1. Authenticating with Spotify...")
    sp = get_spotify_client()
    
    # Get current user info
    user = sp.current_user()
    print(f"   Logged in as: {user['display_name']} ({user['id']})")
    
    # Fetch playlists
    print("\n2. Fetching playlists...")
    playlists = get_user_playlists(sp)
    print(f"   Found {len(playlists)} playlists")
    
    if not playlists:
        print("No playlists found!")
        return
    
    # Process each playlist
    print("\n3. Processing playlists:")
    playlists_data = {}
    
    for idx, playlist in enumerate(playlists, 1):
        playlist_name = playlist['name']
        playlist_id = playlist['id']
        track_count = playlist['tracks']['total']
        
        print(f"   [{idx}/{len(playlists)}] {playlist_name} ({track_count} tracks)...")
        
        tracks = get_playlist_tracks(sp, playlist_id)
        
        if tracks:
            df = pd.DataFrame(tracks)
            sheet_name = sanitize_sheet_name(playlist_name)
            
            # Handle duplicate sheet names
            original_sheet_name = sheet_name
            counter = 1
            while sheet_name in playlists_data:
                sheet_name = f"{original_sheet_name[:28]}_{counter}"
                counter += 1
            
            playlists_data[sheet_name] = df
    
    # Export to Excel
    print("\n4. Exporting to Excel...")
    export_to_excel(playlists_data)
    
    print("\n" + "=" * 50)
    print("Export complete!")


if __name__ == "__main__":
    main()
