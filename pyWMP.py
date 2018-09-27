"""
======================================================================
pyWMP v1.0.0: Provide a standard Python interface to Windows Media Player
======================================================================

See: http://mapletree.wikia.com/wiki/Python/pyWMP

Made to export songs and playlists from Windows Media Player, pyWMP
consists of a set of Python classes based on the win32com package. The
classes abstract the complexity of the COM interface - playlists are
accessed as a dictionary and a group of songs is simply a list of song
elements.

======================================================================

Copyright (c) 2013 Jacob Maples

This work is licensed under the Creative Commons Attribution-ShareAlike
3.0 Unported License. To view a copy of this license, visit
http://creativecommons.org/licenses/by-sa/3.0/.

The above copyright notice and this permission notice shall be
included in all copies or substantial portions of the Software. 

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

======================================================================

Based on the win32com package, pyWMP contains several wrapper classes to
provide a more standard interface to Windows Media Player. Developed on
Windows 7

class pyWMP()
The main class to handle interfacing with Windows Media Player

    .getPlaylists(self, name=None)
    Retrieve all, or one particular, playlist(s). Returns a pyWMPplaylist
    object
    
    .getSongs(self, playlist=None)
    Retrieve all songs, or the songs in a particular playlist. Returns a
    pyWMPsonglist object
    
    .removeLists(self, minSongs=1)
    Removes any playlists that are empty or have fewer than a minimum number
    of songs. Note that this does not delete the actual playlist file, but
    simply removes it from the Windows Media Player library.
    
    .listBroken(self, playlist=None, remove=False)
    List and optionally remove any songs from the Windows Media Player library
    if the song file cannot be found.

class pyWMPplaylist(dict)
A wrapper class to handle multiple playlists

    .exportPlaylists(self, path=None):
    Create a copy of the playlist, pointing to the original song files    

class pyWMPsonglist(list):
A wrapper class to handle a list of songs

    .filterByAttribute(self, attribute='UserRating', test='attribute > 75', label='5-star', keep=True, verbose=True):
    A function to filter a list of songs based on metadata
    
    .filterUnique(self, target_path=None):
    A function to filter a list of songs based on which songs do not exist in
    a particular target directory
    
    .__playlistEntry_M3U__(self, IWMPMedia, URL=''):
    A helper function to return a M3U playlist entry for a particular song
    
    .exportPlaylist(self, filepath=None):
    Create a playlist for the list of songs, pointing to the original song files    
    
    .exportSongs(self, playlist_path='', dest_dir='', source_dir=''):
    Copy a playlist for the list of songs to a specified location and copy any
    songs that do not already exist

Examples of Windows Media Player song attributes are shown below:

Categorizing Properties
5 : AcquisitionTimeYearMonthDay = 8/25/2012
25 : FileType = mp3
30 : MediaType = audio
27 : Is_Protected = False
64 : UserRating = 50

Descriptive Properties
9 : Author = The 71's
6 : AlbumID = We Are The Seventy Ones
51 : Title = Confession
46 : ReleaseDateYearMonthDay = 1/2/2012
75 : WM/Genre = Rock & Roll

23 : Duration = 246.282
24 : FileSize = 9850000
48 : SourceURL = G:\songs\we_are_the_seventy_ones_sample_ep\02_confession.mp3

Other Properties
0 : AcquisitionTime = 8/25/2012 6:33:04 PM
1 : AcquisitionTimeDay = 25
2 : AcquisitionTimeMonth = 8
3 : AcquisitionTimeYear = 2012
4 : AcquisitionTimeYearMonth = 8/25/2012
7 : AlbumIDAlbumArtist = We Are The Seventy Ones*;*
8 : AudioFormat = {00000055-0000-0010-8000-00AA00389B71}
11 : Bitrate = 320000
14 : CanonicalFileType = mp3
18 : DefaultDate = 1/2/2012 12:01:00 AM
19 : Description = 0
20 : DisplayArtist = The 71's
28 : IsVBR = False
41 : ReleaseDate = 1/2/2012 12:01:00 AM
42 : ReleaseDateDay = 2
43 : ReleaseDateMonth = 1
44 : ReleaseDateYear = 2012
45 : ReleaseDateYearMonth = 1/2/2012
47 : RequestState = 2
52 : TrackingID = {572C1079-A876-4CC3-A361-2B5DC45F4135}
55 : UserEffectiveRating = 50
57 : UserPlayCount = 0
58 : UserPlaycountAfternoon = 0
59 : UserPlaycountEvening = 0
60 : UserPlaycountMorning = 0
61 : UserPlaycountNight = 0
62 : UserPlaycountWeekday = 0
63 : UserPlaycountWeekend = 0
65 : UserServiceRating = 0
67 : WM/AlbumTitle = We Are The Seventy Ones
73 : WM/EncodedBy = iTunes 10.6.3
81 : WM/MediaClassPrimaryID = {D1607DBC-E323-4BE2-86A1-48A42A28441E}
82 : WM/MediaClassSecondaryID = {00000000-0000-0000-0000-000000000000}
94 : WM/TrackNumber = 2
100 : WM/Year = 2012
"""

import win32com.client
import os
from shutil import copy2 as filecopy
from glob import glob


def printDuration(seconds=0):
    seconds = int(seconds)
    hr = seconds/3600
    seconds = seconds % 3600
    min = seconds/60
    seconds = seconds % 60
    return '{0:d}:{1:02d}:{2:02d}'.format(hr, min, seconds)


def printSize(bytes=0):
    if bytes < 10000:
        return str(bytes)+' B'
    elif bytes < 1024*10000:
        return str(bytes/1024)+' kB'
    elif bytes < 1024**2*10000:
        return str(bytes/1024**2)+' MB'
    else:
        return str(bytes/1024**3)+' GB'


def printSize(bytes=0):
    if bytes < 1024*1000:
        return '< 1 MB'
    else:
        return str(bytes/1024**2)+' MB'

        
def commonPath(filelist=[]):
    # Return directory prefix used by 50%+ of entries
    prefixlist = []
    checkedlist = set()
    # Order doesn't matter, sets are more efficient
    filelist = set(filelist)
    for f in filelist:
        parts = f.split(os.path.sep)[:-1]
        if len(parts) < len(prefixlist):
            # nothing left to check in this song
            #print '  Not enough parts',parts,'in song',f
            continue

        # check how many files start with...            
        check = os.path.sep.join(parts[0:1+len(prefixlist)])
        if check in checkedlist:
            # This was checked already!
            continue
        
        checkedlist.add(check)
        #print '  Checking',check,'from song',f
        check = [ f2.startswith(check) for f2 in filelist ]
        if sum(check)*2 > len(filelist):
            prefixlist += [ parts[len(prefixlist)] ]
            #print 'Prefix list updated',prefixlist
        else:
            continue
        
    return os.path.sep.join(prefixlist)
    
      
################################################################################

def addMissing(rootDir='', extensions=['wav', 'wma', 'm4a', 'mp3']):
    #Ensure all songs on disk are in the library

    print 'Gathering library information...',
    songs = getSongs()
    path = {}
    for s in songs:
        if not os.access(s.sourceURL, os.F_OK):
            # ignore path if song doesn't exist
            continue
        else: # add path
            tmp = s.sourceURL.split('\\')
            for j in range(1, len(tmp)):
                k = '\\'.join(s.sourceURL.split('\\')[:j])
                if not path.has_key(k):
                    path[k] = []
            path[k] += [s.sourceURL]
    print len(songs),'songs in',len(path.keys()),'directories.'

    added = 0
    if rootDir: # walk the directory to search for all songs
        print 'Searching',rootDir,'for missing songs (','*.'+', *.'.join(extensions),')...'
        for root, dirs, files in os.walk(unicode(rootDir)):
            if path.has_key(root):
                for s in files:
                    f = root+'\\'+s
                    if s.split('.')[-1] not in extensions:
                        continue # not a song
                    elif f in path[root]:
                        continue # song accounted for
                    else: # add song to library
                        print ' ',unicode(f.encode('utf8'))
                        added += 1
                        try:
                            w.mediaCollection.add(f)
                        except:
                            print 'broke on adding ',f
            else: # whole directory missing
                for s in files:
                    f = root+'\\'+s
                    if s.split('.')[-1] not in extensions:
                        continue # not a song
                    else: # add song to library
                        print ' ',unicode(f.encode('utf8'))
                        added += 1
                        w.mediaCollection.add(f)
    else: # check each path for missing songs
        print 'Searching directories for missing songs (','*.'+', *.'.join(extensions),')...'
        for i in path.keys():
            songs = os.listdir(unicode(i))
            if len(songs) == len(path[i]):
                continue # all songs accounted for
            else: # songs don't add up
                for s in songs:
                    f = i+'\\'+s
                    if s.split('.')[-1] not in extensions:
                        continue # not a song
                    elif f in path[i]:
                        continue # song accounted for
                    else: # add song to library
                        print ' ',unicode(f)
                        added += 1
                        w.mediaCollection.add(f)

    print 'Done!',added,'songs added'

# To do: Find songs that are in no playlists

# manually get a list of songs so getSongs works...
w=win32com.client.gencache.EnsureDispatch('WMPlayer.OCX',0)
songs=w.mediaCollection.getByAttribute('MediaType','audio')
songs = [ songs[i] for i in range(songs.count) ]

print "\nStill to encorporate:"
print "addMissing(rootDir='', extensions=['wav', 'wma', 'm4a', 'mp3'])"

################################################################################
class pyWMPsonglist(list):
    """A class to create a more Python-y interface to a list of one or more
Windows Media Player songs

To do:
Add method to save list to WMP playlist (overwrite)
Add method to append list to WMP playlist
"""
    def __init__(self, songlist, listname=None):
        list.__init__(self, songlist)
        #self.__dict__ = {'name':listname}
        self.name = listname


    def __repr__(self):
        time=0
        size=0
        for song in self:
            if len(song.getItemInfo('FileSize')) > 0:
                size += int(song.getItemInfo('FileSize'))
            #else:
            #    print '\n'+pl[i].name+' : '+pl[i][j].sourceURL
            if len(song.getItemInfo('Duration')) > 0:
                time += float(song.getItemInfo('Duration'))

        s = unicode('{0:4d} songs  {1:>9s}  {2:>7s}')
        return s.format(len(self), printDuration(time), printSize(size))


    def listFiles(self):
        """list songs by size, length, and file location"""
        print 'Files in song list "'+str(self.name)+'" :'
        print '{0:>4s} {1:>9s} {2:>6s}  {3}'.format('Idx', 'Duration', 'Size', 'File Location')
        output = []
        for i in range(len(self)):
            size = printSize(int(self[i].getItemInfo('FileSize')))
            time = printDuration(float(self[i].getItemInfo('Duration')))
            print '{0:>3d}) {1:>9s} {2:>6s}  {3}'.format(i, time, size, self[i].sourceURL)
            output += [ self[i].sourceURL ]
        return output


    def describe(self, attrList=['UserRating', 'WM/Genre'], minSongs=1):
        """ list songs by different categories
['UserRating', 'WM/Genre', 'Author', 'WM/AlbumTitle']

Categorizing Properties
5 : AcquisitionTimeYearMonthDay = 8/25/2012
25 : FileType = mp3
30 : MediaType = audio
27 : Is_Protected = False
64 : UserRating = 50

Descriptive Properties
9 : Author = The 71's
6 : AlbumID = We Are The Seventy Ones
51 : Title = Confession
46 : ReleaseDateYearMonthDay = 1/2/2012
75 : WM/Genre = Rock & Roll"""
        for attribute in attrList:
            print '\n--', attribute,
            d = {}
            for song in self:
                attr_set = set([ song.getAttributeName(i) for i in range(song.attributeCount) ])
                if len(attr_set) == 0:
                    print 'Song has no attributes, something bad happened'
                    break
                elif attribute not in attr_set:
                    continue

                value = unicode(song.getItemInfo(attribute))

                if d.has_key(value):
                    # update values
                    d[value]['count'] += 1
                    d[value]['size'] += int(song.getItemInfo('FileSize'))

                    if len(song.getItemInfo('Duration')) > 0:
                        d[value]['time'] += float(song.getItemInfo('Duration'))
                else:
                    # add values
                    d[value] = {'count':1, 'time':0, 'size':int(song.getItemInfo('FileSize'))}

                    if len(song.getItemInfo('Duration')) > 0:
                        d[value]['time'] = float(song.getItemInfo('Duration'))

            width = max([ len(k) for k in d.keys() ])
            print '-'*(width+32-len(attribute))
            for k in sorted(d.keys()):
                if d[k]['count'] < minSongs:
                    continue
                s = unicode('  {0:'+str(width)+'s} : {1:4d} songs  {2:>9s}  {3:>7s}')
                print s.format(k, d[k]['count'], printDuration(d[k]['time']), printSize(d[k]['size']))


    def getAttributes(self, minSongs=2):
        """Build a list of non-trivial attributes for filtering"""
        print "Building attribute histogram for",len(self),"songs...",
        attr = {}
        for song in self:
            for i in range(song.attributeCount):
                k = song.getAttributeName(i)
                v = song.getItemInfo(k)
                if attr.has_key(k):
                    if attr[k].has_key(v):
                        attr[k][v] += 1
                    else:
                        attr[k][v] = 1
                else:
                    attr[k] = {v:1}

        print "Done!\nRemoving trivial attributes...",
        for k in attr.keys():
            # remove attribute if there's one potential value
            if len(attr[k].keys()) == 1:
                del attr[k]
                continue

            # filter values based on the minSongs threshold
            for v in attr[k].keys():
                if attr[k][v] < minSongs:
                    del attr[k][v]

            # remove attribute if there's < 2 potential values
            if len(attr[k].keys()) < 2:
                del attr[k]

        print "Done!"
        return attr
    

    def filterByAttribute(self, attribute='UserRating', test='attribute > 75', label='5-star', keep=True, verbose=True):
        """Filter list of songs to keep or remove songs that match an expression based on 1 attribute
filterByAttribute(self, attribute='UserRating', test='attribute > 75', label='5-star', keep=True, verbose=True)
attribute : Attribute to filter on (some examples shown below)
test      : Expression to evaluate. Must include "attribute", which will be
            replaced with the attribute value for each song.
label     : Label for the resulting list of songs
keep      : Control whether passing the test means a song is kept (True) or removed (False)
verbose   :

Categorizing Properties
5 : AcquisitionTimeYearMonthDay = 8/25/2012
25 : FileType = mp3
30 : MediaType = audio
27 : Is_Protected = False
64 : UserRating = 50

Descriptive Properties
9 : Author = The 71's
6 : AlbumID = We Are The Seventy Ones
51 : Title = Confession
46 : ReleaseDateYearMonthDay = 1/2/2012
75 : WM/Genre = Rock & Roll"""

        if verbose:
            print 'Filtering',self.name,'playlist...'
            print repr(self)
            if keep:
                print 'Keeping',
            else:
                print 'Removing',
            print 'songs for',test.replace('attribute', attribute)

        new_pl = []

        # Filter list of songs by attribute
        for song in self:
            attr_set = set([ song.getAttributeName(i) for i in range(song.attributeCount) ])
            if len(attr_set) == 0:
                print 'Song has no attributes, something bad happened'
                break
            elif attribute not in attr_set:
                if verbose:
                    print ' ',song.sourceURL,'has no attribute',attribute
                continue

            # check if song should be kept
            a = song.getItemInfo(attribute)
            # put extra quotes around result unless it's a number
            if not a.isdigit():
                a = '"' + a + '"'
            check = (eval(test.replace('attribute', a)) == keep)

            if check:
                new_pl += [ song ]
                continue
            else:
                #print ' ',song.sourceURL,a+op+str(value),'is false'
                pass

        if verbose:
            print 'Done! Creating',label,'playlist.'

        pl = pyWMPsonglist(new_pl, label)
        print repr(pl)
        return pl
    

    def filterUnique(self, target_path=None):
        # Filter song list to remove duplicates or songs that exist in given directory

        pathset = set()
        new_pl = []

        if target_path:
            # build recursive file listing of target directory
            # use set since order doesn't matter here
            fileList = {}
            for root, subFolders, files in os.walk(target_path):
                for file in files:
                    if fileList.has_key(file):
                        fileList[file] += [ os.path.join(root,file) ]
                    else:
                        fileList[file] = [ os.path.join(root,file) ]

        for s in self:
            if s.sourceURL in pathset:
                continue
            elif target_path:
                # check if filename is in path
                if fileList.has_key(s.sourceURL.split(os.path.sep)[-1]):
                    continue
                    # if file names match, compare checksums?

            pathset.add(s.sourceURL)
            new_pl += [ s ]

        pl = pyWMPsonglist(new_pl, self.name+' unique')
        print repr(pl)
        return pl


    def __playlistEntry_M3U__(self, IWMPMedia, URL=''):
        # generate a M3U playlist entry for a given song
        s = IWMPMedia
        string = u'#EXTINF:'
        if len(s.getItemInfo('Duration')) > 0:
            string += str(int(float(s.getItemInfo('Duration'))))+','
        else: # assume 2 minutes for other songs
            string += '120,'

        if len(s.getItemInfo('Author')) > 0:
            string += s.getItemInfo('Author')+' - '
        else: # use the directory
            tmp = s.sourceURL.split(os.path.sep)[-2] # get the directory
            tmp = tmp.replace('_',' ') # replace underscores with spaces
            tmp = tmp.title() # capitalize the first letter in each word
            string += tmp+' - '

        if len(s.getItemInfo('Title')) > 0:
            string += s.getItemInfo('Title')
        else: # use the filename
            tmp = s.sourceURL.split(os.path.sep)[-1] # get the filename
            tmp = tmp.split('.')[0] # drop the extension
            tmp = tmp.replace('_',' ') # replace underscores with spaces
            tmp = tmp.title() # capitalize the first letter in each word
            string += tmp

        if URL: # use specified URL
            string += '\n'+URL+'\n'
        else: # use URL from library
            string += '\n'+s.sourceURL+'\n'

        return string
    

    def exportPlaylist(self, filepath=None):
        # generate playlist for given songs in library
        # placeholder for doing multiple formats
        mode = 'M3U'

        # use the current directory and list name by default
        if filepath == None:
            filepath = os.path.join(os.getcwd(),self.name)

        # placeholder for doing multiple formats
        filepath = filepath+'.m3u'
        header = '#EXTM3U\n'
        playlistEntry = self.__playlistEntry_M3U__

        print 'Creating '+mode+' playlist for "'+self.name+'" songlist at '+filepath,
        f = file(filepath, 'w')
        f.write(header)
        print '...',

        for s in self:
            string = playlistEntry(s)
            try:
                f.write(string.encode('utf8'))
            except:
                print u'Failed on '+repr(string)
        f.close()
        print 'Done!'
        

    def exportSongs(self, playlist_path='', dest_dir='', source_dir=''):
        
        if not source_dir: # Default source directory is a common path
            source_dir = commonPath([ f.sourceURL for f in self ])

        if not dest_dir: # Default destination is playlist location
            if playlist_path:
                dest_dir = os.path.sep.join(playlist_path.split(os.path.sep)[:-1])
            else:
                dest_dir = os.path.sep.join([os.getcwd(),'SongExport'])
            
        if not playlist_path:
            playlist_path = os.path.sep.join([dest_dir, self.name])
        
        if not os.path.exists(dest_dir):
            # A recursive directory would be more robust
            os.makedirs(dest_dir)

        # placeholder for doing multiple formats
        mode = 'M3U'

        # placeholder for doing multiple formats
        playlist_path = playlist_path+'.m3u'
        header = '#EXTM3U\n'
        playlistEntry = self.__playlistEntry_M3U__

        print 'Exporting '+mode+' playlist for "'+self.name+'" songlist at '+playlist_path
        print '  Exporting songs to '+dest_dir,
        f = file(playlist_path, 'w')
        f.write(header)
        print '...',

        # Determine which songs will be copied over
        pl = self.filterUnique(dest_dir)

        for s in pl:
            #Make sure song exists...
            if not os.access(s.sourceURL, os.F_OK):
                continue

            if s.sourceURL.startswith(source_dir):
                # copy releative path to new location
                path = s.sourceURL[len(source_dir)+1:]
            else: # copy full path
                path = s.sourceURL.replace(':','')

            # check if the song has already been copied
            if os.access(os.path.sep.join([dest_dir, path]), os.F_OK):
                print 'Duplicate song!!', path
                continue

            # strip file name from path and add leading slash
            path = os.path.sep.join([dest_dir]+path.split(os.path.sep)[:-1])

            # check if path exists
            if not os.access(path, os.F_OK):
                os.makedirs(path) # create directory if needed

            # copy the file
            filecopy(s.sourceURL, path)

        for s in self:        
            # Get the playlist entry for each song
            string = playlistEntry(s).replace(source_dir+os.path.sep, '')
            try:
                f.write(string.encode('utf8'))
            except:
                print u'Failed on '+repr(string)                
        f.close()
        print 'Done!'
            
        pass

################################################################################

class pyWMPplaylist(dict):
    """A class to create a more Python-y interface to a set of one or more
Windows Media Player playlist

To do:
When playlist is changed, update Windows Media Player"""
    def __init__(self, win32com_ptr, name=None):
        #def __init__(self, win32com_ptr, name=None):
        # get playlists
        if name:
            pl=win32com_ptr.playlistCollection.getByName(name)
        else:
            pl=win32com_ptr.playlistCollection.getAll()

        # check if there's anything to return
        if pl.count == 0:
            return None

        for i in range(pl.count):
            if pl[i].count == 0:
                self[pl[i].name] = pyWMPsonglist([], pl[i].name)
            else:
                self[pl[i].name] = pyWMPsonglist([ pl[i][j] for j in range(pl[i].count) ], pl[i].name)

        self.name = name

        # save the handle for Windows Media Player
        self.wmp = win32com_ptr

    def __delitem__(self, key):
        """If a playlist is deleted, remove it from Windows Media Player"""
        pl=self.wmp.playlistCollection.getByName(key)

        # make sure it's a valid playlist
        if pl.count > 0:
            # remove playlist from Windows Media Player
            # Note that the playlist file will be left on disk
            self.wmp.playlistCollection.remove(pl[0])

            # Regenerate the playlists since deleting the key causes issues
            self = self.__init__(self.wmp, self.name)

    def __repr__(self):
        width = max([ len(k) for k in self.keys() ])
        out = '-- Playlists '+'-'*(width+22)+'\n'
        s = unicode(' {0:'+str(width)+'s} : ')
        for k in self.keys():
            out += s.format(k)+repr(self[k])+'\n'
        return out

    def exportPlaylists(self, path=None):
        # use the current directory and list name by default
        if path == None:
            path = os.getcwd()+os.path.sep

        # export each playlist
        for k in self.keys():
            self[k].exportPlaylist(path+k)

################################################################################

class pyWMP:
    """A class to create a more Python-y interface to Windows Media Player"""

    def __init__(self):
        # Get a pointer for Windows Media Player
        self.wmp = win32com.client.gencache.EnsureDispatch('WMPlayer.OCX',0)

    def getPlaylists(self, name=None):
        if name:
            print 'Getting songs for "'+name+'" playlist...',
        else:
            print 'Getting songs for all playlists...',
        pl = pyWMPplaylist(self.wmp, name)
        print 'Done!\n'+repr(pl)
        return pl

    def getSongs(self, playlist=None):
        if playlist:
            print 'Getting songs for "'+playlist+'" playlist...',
            pl = pyWMPplaylist(self.wmp, playlist)
            if pl: # make sure playlist is valid
                pl = pl[playlist]
            else:
                pl = pyWMPsonglist([], 'playlist')
        else:
            print 'Getting all songs...',
            pl=self.wmp.mediaCollection.getByAttribute('MediaType','audio')
            pl=pyWMPsonglist([ pl[i] for i in range(pl.count) ], 'All Music')

        print 'Done!\n'+repr(pl)
        return pl

    def removeLists(self, minSongs=1):
        """List playlists and remove any that are empty (or less than the minimum number of songs specified)"""
        pl = pyWMPplaylist(self.wmp)

        empty = [ (k, len(pl[k])) for k in pl.keys() ]
        empty = [ i for i in empty if i[1] < minSongs ]

        print "Deleting playlists:",

        for i in empty:
            if i[1] == 1:
                print "{0} ({1:d} song),".format(i[0], i[1]),
            else:
                print "{0} ({1:d} songs),".format(i[0], i[1]),

            del pl[i[0]]

    def listBroken(self, playlist=None, remove=False):
        #Check library for songs that don't exist

        if playlist:
            # search given playlist
            print 'Searching playlist "'+playlist+'" for missing files...',
            pl = pyWMPplaylist(self.wmp, playlist)
        else: # search all songs in library
            print 'Searching for missing files...',
            pl=self.getSongs()

        bogus = []
        songs = self.getSongs(playlist)
        for i,s in enumerate(songs):
            if not os.access(s.sourceURL, os.F_OK):
                # remember song index if for bogus tracks
                bogus += [ i ]

        if remove:
            # remove bogus songs in reverse order
            print 'Removing',len(bogus),'files:'
            bogus.reverse()
            if playlist:
                for i in bogus:
                    print ' ',songs[i].sourceURL
                    pl[0].removeItem(pl[0][i])
            else:
                for i in bogus:
                    print ' ',songs[i].sourceURL
                    w.mediaCollection.remove(songs[i], False)
        else: # just show the results
            print 'Done!'
            for i in bogus:
                print ' ',songs[i].sourceURL

################################################################################

if __name__ == '__main__':
    wmp = pyWMP()
