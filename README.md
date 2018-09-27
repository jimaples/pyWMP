# pyWMP

pyWMP is licensed under a [GPL-3.0 license](https://github.com/jimaples/pyWMP/blob/master/LICENSE)

## Background

I wanted to copy songs from Windows Media Player (WMP) on my PC to my phone without having to manually create and organize playlists on the phone. After the previous win32com presentation, I realized that Python should be able to access WMP and do what I needed.

## Functionality

To port songs and playlists from WMP to a phone, certain functionality is necessary
* Retrieve selected (or all) playlists from WMP
* Retrieve songs from WMP
* Copy a group of songs and a playlist referencing those songs to a specified location
  * Playlist must be in a format the phone can recognize
  * Playlist must use relative paths sInce phone file structure will not match PC
* Filter a group of songs based on attribute values (e.g. user rating)
* List the attributes which may be useful for filtering and potential values
Other functionality is not strictly needed, but improves the useability
* List and possibly delete empty or throwaway playlists from WMP library
* List and possibly delete invalid files (e.g. file not found) from WMP library
* Export a playlist for a group of songs (without copying or relative paths, useful for sharing playlists on a home network)
* Remove duplicates from a group of songs
* Filter a group of songs based on which files do not exist in a specified directory (cuts down on copy time)
* List the number of songs, total play time, and total file size for a group of songs
* Describe a group of songs based on selected attributes

## Windows Media Player COM Interface

There is not much online for the WMP interface, but using Pythonwin, I was able to investigate how it worked. In the first iteration of this project, I developed a number of standalone scripts to provide the necessary functionality. To abstract out the COM interface, I decided to convert these scripts to the object-oriented approach described in the pyWMP section.

WMP is organized around playlists and media, seperately. The following code snippet shows how the interface is initiated, how to retrieve all playlists and audio files, and the objects used.
```python
>>> w=win32com.client.gencache.EnsureDispatch('WMPlayer.OCX',0)
>>> w
<win32com.gen_py.Windows Media Player.IWMPPlayer4 instance at 0x32240664>
>>> songs=w.mediaCollection.getByAttribute('MediaType','audio')
>>> songs[0]
<win32com.gen_py.Windows Media Player.IWMPMedia instance at 0x32242184>
>>> pl=w.playlistCollection.getAll()
>>> pl[0]
<win32com.gen_py.Windows Media Player.IWMPPlaylist instance at 0x33972832>
Both playlists and media have sets of attributes that are accessed by name. A value indicates how many attributes an object has and a method is used to retrieve the attribute name. The following code snippit shows interactions with playlists:
>>> pl[0].attributeCount
4
>>> pl[0].attributeName(3)
u'Title'
>>> pl[0].getItemInfo('Title')
u'My Playlist'
>>> pl[0].count # lookup how many songs this playlist contains
144
>>> # Access playlist directly 
>>> w.playlistCollection.getByName('My Playlist')[0].count
144
```

The following code snippit shows interactions with songs. Note that the file location is available directly (and that songs can have a lot of attributes).
```python
>>> pl[0][0] # first song in playlist 0
<win32com.gen_py.Windows Media Player.IWMPMedia instance at 0x33974232>
>>> pl[0][0].attributeCount
100
>>> pl[0][0].getAttributeName(25)
u'FileType'
>>> pl[0][0].getItemInfo(25) # Bad attribute requests return empty strings
u
>>> pl[0][0].getItemInfo('FileType')
u'mp3'
>>> pl[0][0].sourceURL # file location
u'G:\\songs\\Gotan Project\\05 Santa Maria (del Buen Ayre).mp3'
>>> # Access song directly
>>> f = u'G:\\songs\\Gotan Project\\05 Santa Maria (del Buen Ayre).mp3'
>>> w.mediaCollection.getByAttribute('SourceURL', f)[0]
<win32com.gen_py.Windows Media Player.IWMPMedia instance at 0x33974352>
>>> # You can also getByAlbum/Author/Genre/Name
>>> w.mediaCollection.getByAuthor('Gotan Project').count
14
```

## pyWMP

Clearly, the COM interface doesn't line up very well with the standard Python data structures (e.g. dictionaries, lists). To access WMP from Python without being a COM expert, I decided to create a wrapper class based on standard data structures. I created 3 wrapper classes to this end:
* `pyWMP` : The WMP interface itself, controlled by methods
* `pyWMPplaylist` : Treat playlists as a dictionary. Playlist names are used as keys. The values are pyWMPsonglist entities. Methods of the pyWMPplaylist rely on the pyWMPsonglist class
* `pyWMPsonglist` : Treat a group of songs as a list. Elements of the list are win32com.gen_py.Windows Media Player.IWMPMedia objects

Python libraries used:
* `win32com.client` (COM access to Windows Media Player)
* `os` (Check file existence, walk directories, cross-platform path handling)
* `shutil.copy2` (File copy)

### Issues Encountered
It took a few iterations to get the pyWMPsonglist to initialize without crashing Python
* Problem: Although user-defined classes can always have attributes, the list has no `__dict__` parameter to hold user-defined attributes.
```python
def __init__(self, songlist, listname=None):
    self = songlist
    self.name = listname # crashes script
    self.__dict__ = {'name':listname} # still crashes
```
* Fix: Instead of simply setting self with the list input, use the list data type's `__init__` method as shown below
```python
def __init__(self, songlist, listname=None):
        list.__init__(self, songlist)
        self.name = listname
```
Showing the representation of playlists with hundreds or thousands of songs caused some unwanted delays and was not very helpful since the songs just show the object type and memory address.
Problem: Default representation (from the output of a method or displayed manually) for a song list was unwieldy.
```python
{'My Playlist':[<win32com.gen_py.Windows Media Player.IWMPMedia instance at 0x32
242184>, <win32com.gen_py.Windows Media Player.IWMPMedia instance at 0x32242185>
, <win32com.gen_py.Windows Media Player.IWMPMedia instance at 0x32242186>, ...hu
ndreds more..., <win32com.gen_py.Windows Media Player.IWMPMedia instance at 0x32
242184>], 'Playlist2':[<win32com.gen_py.Windows Media Player.IWMPMedia instance 
at 0x32242174>, <win32com.gen_py.Windows Media Player.IWMPMedia instance at 0x32
242175>, <win32com.gen_py.Windows Media Player.IWMPMedia instance at 0x32242176>
, ...hundreds more..., <win32com.gen_py.Windows Media Player.IWMPMedia instance 
at 0x32242174>]
}
```
* Fix: Add a `__repr__` methods to pyWMPsonglist (to display the number of songs, total time, and total file size) and pyWMPplaylist (to display playlist name, format the representation, and call the pyWMPsonglist `__repr__` method).
```
-- Playlists -----------------------------
My Playlist  103 songs   5:12:43  1027 MB
Playlist2    341 songs  18:41:27  4371 MB
```

### Methods and Attributes

`class pyWMP()` 
The main class to handle interfacing with Windows Media Player
* `.getPlaylists(self, name=None)`
  * Retrieve all, or one particular, playlist(s). Returns a pyWMPplaylist object
* `.getSongs(self, playlist=None)`
  * Retrieve all songs, or the songs in a particular playlist. Returns a pyWMPsonglist object
* `.removeLists(self, minSongs=1)`
  * Removes any playlists that are empty or have fewer than a minimum number of songs. Note that this does not delete the actual playlist file, but simply removes it from the Windows Media Player library.
* `.listBroken(self, playlist=None, remove=False)`
  * List and optionally remove any songs from the Windows Media Player library if the song file cannot be found.

`class pyWMPplaylist(dict)`
A wrapper class to handle multiple playlists
* `.exportPlaylists(self, path=None)`
  * Create a copy of the playlist, pointing to the original song files

`class pyWMPsonglist(list)`
A wrapper class to handle a list of songs
* `.filterByAttribute(self, attribute='UserRating', test='attribute > 75', label='5-star', keep=True, verbose=True)`
  * A function to filter a list of songs based on metadata
* `.filterUnique(self, target_path=None)`
  * A function to filter a list of songs based on which songs do not exist in a particular target directory
* `.__playlistEntry_M3U__(self, IWMPMedia, URL='')`
  * A helper function to return a M3U playlist entry for a particular song
* `.exportPlaylist(self, filepath=None)`
  * Create a playlist for the list of songs, pointing to the original song files    
* `.exportSongs(self, playlist_path='', dest_dir=, source_dir=)`
  * Copy a playlist for the list of songs to a specified location and copy any songs that do not already exist

### Potential Future Upgrades

Although the file does what I need, I can see a few odds and ends that would be useful.
* Open WMP with the songs in a pyWMPsonglist object
* Create a WMP playlist by adding a key/value pair to a pyWMPplaylist object
* Modify the WMP playlist when changing the a value in a pyWMPplaylist object
* Support a playlist format that Android can remove songs from without having to delete the song (may not be possible currently)
* Support other playlist formats
