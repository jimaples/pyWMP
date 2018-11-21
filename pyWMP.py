# noinspection SpellCheckingInspection
"""
======================================================================
pyWMP: Provide a standard Python interface to Windows Media Player
======================================================================

See: https://github.com/jimaples/pyWMP

Made to export songs and playlists from Windows Media Player, pyWMP
consists of a set of Python classes based on the win32com package. The
classes abstract the complexity of the COM interface - playlists are
accessed as a dictionary and a group of songs is simply a list of song
elements.

======================================================================

Based on the win32com package, pyWMP contains several wrapper classes to
provide a more standard interface to Windows Media Player. Developed on
Windows 10

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

# noinspection PyPep8Naming
from win32com.client.gencache import EnsureDispatch as win32com_dispatch
import os
from datetime import datetime as dt
from shutil import copy2 as filecopy
from collections import Counter

eval_globals = {}


# noinspection PyPep8Naming
class pyWMPsonglist(list):
    """A class to create a more Python-y interface to a list of one or more
Windows Media Player songs

To do:
Add method to save list to WMP playlist (overwrite)
Add method to append list to WMP playlist
"""

    def __init__(self, songlist, listname=None):
        super(pyWMPsonglist, self).__init__(songlist)
        self.name = listname

    @staticmethod
    def print_duration(seconds=0.0):
        seconds = int(seconds)
        hr = seconds // 3600
        seconds %= 3600
        minutes = seconds // 60
        seconds %= 60
        return '{0:d}:{1:02d}:{2:02d}'.format(hr, minutes, seconds)

    @staticmethod
    def print_size(num_bytes=0):
        i = 1024 ** 2
        if num_bytes < i:
            return '  < 1 MB'
        else:
            return '{:5d} MB'.format(num_bytes // i)

    def __repr__(self):
        time = 0
        size = 0
        for song in self:
            if len(song.getItemInfo('FileSize')) > 0:
                size += int(song.getItemInfo('FileSize'))
            # else:
            #    print('\n'+pl[i].name+' : '+pl[i][j].sourceURL)
            if len(song.getItemInfo('Duration')) > 0:
                time += float(song.getItemInfo('Duration'))

        s = '{0:4d} songs  {1:>9s}  {2:>7s}'
        return s.format(len(self), self.print_duration(time), self.print_size(size))

    def __contains__(self, item):
        songs = set(i.sourceURL for i in self)
        if type(item) == pyWMPsonglist:
            check = set(i.sourceURL for i in item)
            return not songs.difference(check)
        elif 'IWMPMedia' in item.__doc__:
            return item.sourceURL in songs
        else:
            raise NotImplemented("Not sure...")

    def list_files(self):
        """list songs by size, length, and file location"""
        print('Files in song list "' + str(self.name) + '" :')
        print('{0:>4s} {1:>9s} {2:>6s}  {3}'.format('Idx', 'Duration', 'Size', 'File Location'))
        output = []
        for i in range(len(self)):
            size = self.print_size(int(self[i].getItemInfo('FileSize')))
            time = self.print_duration(float(self[i].getItemInfo('Duration')))
            print('{0:>3d}) {1:>9s} {2:>6s}  {3}'.format(i, time, size, self[i].sourceURL))
            output.append(self[i].sourceURL)
        return output

    def describe(self, attr_list=('UserRating', 'WM/Genre'), min_songs=1):
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
        for attribute in attr_list:
            print('\n--', attribute),
            d = {}
            for song in self:
                attr_set = set([song.getAttributeName(i) for i in range(song.attributeCount)])
                if len(attr_set) == 0:
                    print('Song has no attributes, something bad happened')
                    break
                elif attribute not in attr_set:
                    continue

                value = song.getItemInfo(attribute)

                if value in d:
                    # update values
                    d[value]['count'] += 1
                    d[value]['size'] += int(song.getItemInfo('FileSize'))

                    if len(song.getItemInfo('Duration')) > 0:
                        d[value]['time'] += float(song.getItemInfo('Duration'))
                else:
                    # add values
                    d[value] = {'count': 1, 'time': 0.0, 'size': int(song.getItemInfo('FileSize'))}

                    if len(song.getItemInfo('Duration')) > 0:
                        d[value]['time'] = float(song.getItemInfo('Duration'))

            width = max([len(k) for k in d.keys()])
            print('-' * (width + 32 - len(attribute)))
            for k in sorted(d.keys()):
                if d[k]['count'] < min_songs:
                    continue
                s = '  {0:' + str(width) + 's} : {1:4d} songs  {2:>9s}  {3:>7s}'
                print(s.format(k, d[k]['count'], self.print_duration(d[k]['time']), self.print_size(d[k]['size'])))

    def get_attributes(self, min_songs=2):
        """Build a list of non-trivial attributes for filtering"""
        print("Building attribute histogram for", len(self), "songs..."),
        attr = {}
        for song in self:
            for i in range(song.attributeCount):
                k = song.getAttributeName(i)
                v = song.getItemInfo(k)
                if k in attr:
                    if v in attr[k]:
                        attr[k][v] += 1
                    else:
                        attr[k][v] = 1
                else:
                    attr[k] = {v: 1}

        print("Done!\nRemoving trivial attributes..."),
        for k in attr.keys():
            # remove attribute if there's one potential value
            if len(attr[k].keys()) == 1:
                del attr[k]
                continue

            # filter values based on the min_songs threshold
            for v in attr[k].keys():
                if attr[k][v] < min_songs:
                    del attr[k][v]

            # remove attribute if there's < 2 potential values
            if len(attr[k].keys()) < 2:
                del attr[k]

        print("Done!")
        return attr

    def filter_by_attribute(self, attribute='UserRating', test='attribute > 75',
                            label='5-star', keep=True, verbose=True):
        """Filter list of songs to keep or remove songs that match an expression based on 1 attribute
filter_by_attribute(self, attribute='UserRating', test='attribute > 75', label='5-star', keep=True, verbose=True)
attribute : Attribute to filter on (some examples shown below)
test      : Function or expression to evaluate. Boolean result required.

            Expressions must include "attribute", which will be replaced with
            the attribute value for each song.

            Functions can be passed in for more complicated evaluations or if
            calls to additional libraries or functions are required. The first
            argument of the function must be the attribute value.

label     : Label for the resulting list of songs
keep      : Control whether passing the test means a song is kept (True) or removed (False)
verbose   : Output detailed information for each song

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
            s = 'Filtering ' + self.name + ' playlist... ' + repr(self)
            if keep:
                s += ' Keeping'
            else:
                s += ' Removing'
            s += ' songs for'
            if type(test) == str:
                s += test.replace('attribute', attribute)
            else:
                s += '{0:s}({1:s}) == True'.format(test.__name__, attribute)
            print(s)

        new_pl = []

        # Filter list of songs by attribute
        for song in self:
            attr_set = set([song.getAttributeName(i) for i in range(song.attributeCount)])
            if len(attr_set) == 0:
                print('Song has no attributes, something bad happened')
                break
            elif attribute not in attr_set:
                if verbose:
                    print(' ', song.sourceURL, 'has no attribute', attribute)
                continue

            # check if song should be kept
            a = song.getItemInfo(attribute)
            try:
                if type(test) == str:
                    # put extra quotes around result unless it's a number
                    if not a.isdigit():
                        a = '"' + a + '"'
                    check = (eval(test.replace('attribute', a), eval_globals) == keep)
                elif type(test) == function:
                    # noinspection PyCallingNonCallable
                    check = (test(a) == keep)
                else:
                    raise TypeError('Input "test" must be a string evaluation or function!')

            except NameError as e:
                print(song.getItemInfo('SourceURL'))
                print('  test:', test.replace('attribute', a))
                print('  globals:', globals().keys())
                print(' ', e)
                break

            if check:
                new_pl.append(song)
                continue
            else:
                # print(' ',song.sourceURL,a+op+str(value),'is false')
                pass

        if verbose:
            print('Done! Creating', label, 'playlist.')

        pl = pyWMPsonglist(new_pl, label)
        print(repr(pl))
        return pl

    @staticmethod
    def _isRecent(attribute, days=90):
        """Helper function to see if a timestamp string falls in the window"""
        age = dt.now() - dt.strptime(attribute, "%m/%d/%Y")
        return age.days < days

    # noinspection PyTypeChecker
    def filter_recent(self, days=90, keep=True, verbose=True):
        """Filter songs added recently"""
        s = 'AcquisitionTimeYearMonthDay'
        pl = self.filter_by_attribute(s, lambda a: self._isRecent(a, days), 'New Songs', keep, verbose)
        return pl

    def filter_unique(self, target_path=None, reverse=False):
        """Filter song list to remove duplicates or songs that exist in given directory"""

        pathset = set()
        new_pl = []
        file_list = None

        if target_path:
            # build recursive file listing of target directory
            # use set since order doesn't matter here
            file_list = {}
            for root, subFolders, files in os.walk(target_path):
                for f in files:
                    if f in file_list:
                        file_list[f].append(os.path.join(root, f))
                    else:
                        file_list[f] = [os.path.join(root, f)]

        for s in self:
            if s.sourceURL in pathset:
                continue
            elif file_list:
                check = s.sourceURL.split(os.path.sep)[-1]
                # check if filename is in path
                if check in file_list:
                    if reverse:
                        pass  # target has file, remove
                    else:
                        continue  # target has file, skip
                elif reverse:
                    continue  # target missing file, skip
                    # else, target missing file, add

            pathset.add(s.sourceURL)
            new_pl.append(s)

        pl = pyWMPsonglist(new_pl, self.name + ' unique')
        print(repr(pl))
        return pl

    @staticmethod
    def __playlistEntry_M3U__(iwmp_media, url=''):
        # generate a M3U playlist entry for a given song
        s = iwmp_media
        string = u'#EXTINF:'
        if len(s.getItemInfo('Duration')) > 0:
            string += str(int(float(s.getItemInfo('Duration')))) + ','
        else:  # assume 2 minutes for other songs
            string += '120,'

        if len(s.getItemInfo('Author')) > 0:
            string += s.getItemInfo('Author') + ' - '
        else:  # use the directory
            tmp = s.sourceURL.split(os.path.sep)[-2]  # get the directory
            tmp = tmp.replace('_', ' ')  # replace underscores with spaces
            tmp = tmp.title()  # capitalize the first letter in each word
            string += tmp + ' - '

        if len(s.getItemInfo('Title')) > 0:
            string += s.getItemInfo('Title')
        else:  # use the filename
            tmp = s.sourceURL.split(os.path.sep)[-1]  # get the filename
            tmp = tmp.split('.')[0]  # drop the extension
            tmp = tmp.replace('_', ' ')  # replace underscores with spaces
            tmp = tmp.title()  # capitalize the first letter in each word
            string += tmp

        if url:  # use specified url
            string += '\n' + url + '\n'
        else:  # use url from library
            string += '\n' + s.sourceURL + '\n'

        return string

    def export_playlist(self, playlist_path=None, source_dir=None, dest_dir=None):
        # generate playlist for given songs in library

        # placeholder for doing multiple formats
        mode = 'M3U'

        # use the current directory and list name by default
        if playlist_path is None:
            playlist_path = os.path.join(os.getcwd(), self.name)

        # Keep a log of any errors
        f_err = playlist_path+'_errors.log'
        if os.access(f_err, os.F_OK):
            os.remove(f_err)  # fresh log for each run

        # placeholder for doing multiple formats
        playlist_path += '.m3u'
        header = '#EXTM3U\n'
        playlist_entry = self.__playlistEntry_M3U__

        # print(' dest ' + repr(dest_dir) + ' source ' + repr(source_dir) + repr(dest_dir and source_dir))
        with open(playlist_path, 'w') as f:
            print('Creating ' + mode + ' playlist for "' + self.name + '" songlist at ' + f.name),
            f.write(header)
            print('...'),

            cnt = 0
            for s in self:
                string = playlist_entry(s)

                # Map source to destination directories
                if bool(source_dir) and bool(dest_dir):
                    string = string.replace(source_dir, dest_dir)

                try:
                    f.write(string.encode('utf8'))
                except Exception as e:
                    with open(f_err, 'a') as log:
                        log.write('Failed on ' + repr(string)+'\n')
                        log.write('          ' + repr(e)+'\n')
                    cnt += 1

        if cnt:
            print('{:d} errors, See log for details.'.format(cnt))
        else:
            print('Done!')

        return cnt

    @staticmethod
    def common_path(filelist=None):
        # Return directory prefix used by 50%+ of entries
        if filelist is None:
            filelist = []
        prefixlist = []
        checkedlist = set()
        # Order doesn't matter, sets are more efficient
        filelist = set(filelist)
        for f in filelist:
            parts = f.split(os.path.sep)[:-1]
            if len(parts) < len(prefixlist):
                # nothing left to check in this song
                # print('  Not enough parts',parts,'in song',f)
                continue

            # check how many files start with...
            check = os.path.sep.join(parts[0:1 + len(prefixlist)])
            if check in checkedlist:
                # This was checked already!
                continue

            checkedlist.add(check)
            # print('  Checking',check,'from song',f)
            check = [f2.startswith(check) for f2 in filelist]
            if sum(check) * 2 > len(filelist):
                prefixlist.append(parts[len(prefixlist)])
                # print('Prefix list updated',prefixlist)
            else:
                continue

        return os.path.sep.join(prefixlist)

    def export_songs(self, playlist_path='', dest_dir='', source_dir=''):
        # Copy songs to a given directory

        if not source_dir:  # Default source directory is a common path
            source_dir = self.common_path([f.sourceURL for f in self])

        if not dest_dir:  # Default destination is playlist location
            if playlist_path:
                dest_dir = os.path.sep.join(playlist_path.split(os.path.sep)[:-1])
            else:
                dest_dir = os.path.sep.join([os.getcwd(), 'SongExport'])

        if not playlist_path:
            playlist_path = os.path.sep.join([dest_dir, self.name])

        if not os.path.exists(dest_dir):
            # A recursive directory would be more robust
            os.makedirs(dest_dir)

        # placeholder for doing multiple formats
        mode = 'M3U'

        # placeholder for doing multiple formats
        playlist_path += '.m3u'
        header = '#EXTM3U\n'
        playlist_entry = self.__playlistEntry_M3U__

        print('Exporting ' + mode + ' playlist for "' + self.name + '" songlist at ' + playlist_path)
        print('  Exporting songs to ' + dest_dir),
        with open(playlist_path, 'w') as f:
            f.write(header)
            print('...'),

            # Determine which songs will be copied over
            pl = self.filter_unique(dest_dir)

            for s in pl:
                # Make sure song exists...
                if not os.access(s.sourceURL, os.F_OK):
                    continue

                if s.sourceURL.startswith(source_dir):
                    # copy relative path to new location
                    path = s.sourceURL[len(source_dir) + 1:]
                else:  # copy full path
                    path = s.sourceURL.replace(':', '')

                # check if the song has already been copied
                if os.access(os.path.sep.join([dest_dir, path]), os.F_OK):
                    print('Duplicate song!!', path)
                    continue

                # strip file name from path and add leading slash
                path = os.path.sep.join([dest_dir] + path.split(os.path.sep)[:-1])

                # check if path exists
                if not os.access(path, os.F_OK):
                    os.makedirs(path)  # create directory if needed

                # copy the file
                filecopy(s.sourceURL, path)

            # don't put a song in a playlist more than once
            songs = pyWMPsonglist([])
            for s in self:
                if s in songs:
                    continue
                else:
                    songs.append(s)

                # Get the playlist entry for each song
                string = playlist_entry(s).replace(source_dir + os.path.sep, '')
                try:
                    f.write(string)  # .encode('utf8'))
                except Exception as e:
                    print(u'Failed on ' + repr(string))
                    print(e)

        print('Done!')

    def remove_songs(self, playlist_path='', dest_dir='', source_dir=''):
        # Remove songs from a given directory

        if not source_dir:  # Default source directory is a common path
            source_dir = self.common_path([f.sourceURL for f in self])

        if not dest_dir:  # Default destination is playlist location
            raise TypeError('Must specify what directory to remove songs from!')
        elif not os.path.exists(dest_dir):
            raise TypeError('Specified directory does not exist!')

        if not playlist_path:
            playlist_path = os.path.sep.join([dest_dir, self.name])

        # placeholder for doing multiple formats
        playlist_path += '.m3u'
        if os.path.exists(playlist_path):
            print('Deleting playlist for "' + self.name + '" songlist at ' + playlist_path)
            os.remove(playlist_path)

        print('Deleting "' + self.name + '" songs from ' + dest_dir)

        # Determine which songs will be deleted
        pl = self.filter_unique(dest_dir, reverse=True)

        for s in pl:
            if s.sourceURL.startswith(source_dir):
                # copy relative path to new location
                path = s.sourceURL[len(source_dir) + 1:]
            else:  # copy full path
                path = s.sourceURL.replace(':', '')

            # strip file name from path and add leading slash
            path = os.path.sep.join([dest_dir] + path.split(os.path.sep)[:-1])

            # check if file exists
            if not os.access(path, os.F_OK):
                continue
            else:
                os.remove(path)
        print('Done!')

    def remove_duplicates(self):
        # Check playlist for duplicate songs
        raise NotImplemented("TODO: Check playlist for duplicate songs")

        # noinspection PyUnreachableCode
        song_counts = Counter([s.getItemInfo('SourceURL') for s in self])

        # check for songs that appear more than once
        for k, v in paths.iteritems():
            if v > 1:
                pass
                test = pl.wmp.playlistCollection.getByName('World')[0]
                test.removeItem(test.Item(50))


################################################################################

# noinspection PyPep8Naming
class pyWMPplaylist(dict):
    """A class to create a more Python-y interface to a set of one or more
Windows Media Player playlist

To do:
When playlist is changed, update Windows Media Player"""

    def __init__(self, win32com_ptr, name=None, **kwargs):
        super(pyWMPplaylist, self).__init__(**kwargs)

        # get playlists
        if name:
            pl = win32com_ptr.playlistCollection.getByName(name)
        else:
            pl = win32com_ptr.playlistCollection.getAll()

        # check if there's anything to return
        for i in range(pl.count):
            if pl[i].count == 0:
                self[pl[i].name] = pyWMPsonglist([], pl[i].name)
            else:
                self[pl[i].name] = pyWMPsonglist([pl[i][j] for j in range(pl[i].count)], pl[i].name)

        self.name = name

        # save the handle for Windows Media Player
        self.wmp = win32com_ptr

    def __delitem__(self, key):
        """If a playlist is deleted, remove it from Windows Media Player"""
        pl = self.wmp.playlistCollection.getByName(key)

        # make sure it's a valid playlist
        if pl.count > 0:
            # remove playlist from Windows Media Player
            # Note that the playlist file will be left on disk
            self.wmp.playlistCollection.remove(pl[0])

            # Regenerate the playlists since deleting the key causes issues
            # noinspection PyMethodFirstArgAssignment
            self = self.__init__(self.wmp, self.name)

    def __repr__(self):
        width = max([len(k) for k in self.keys()])
        out = '-- Playlists ' + '-' * (width + 22) + '\n'
        s = ' {0:' + str(width) + 's} : '
        for k in self.keys():
            out += s.format(k) + repr(self[k]) + '\n'
        return out

    def export_playlists(self, playlist_path=None, source_dir=None, dest_dir=None):
        # use the current directory and list name by default
        if playlist_path is None:
            playlist_path = os.getcwd() + os.path.sep

        # export each playlist
        errors = 0
        s = ''
        for k in self.keys():
            if len(self[k]) == 0:
                continue  # skip empty playlists
            elif k in ('All Music', 'All Pictures', 'All Video',
                       'Music auto rated at 5 stars', 'Pictures rated 4 or 5 stars'):
                continue  # skip WMP default playlists

            cnt = self[k].export_playlist(playlist_path + k, source_dir, dest_dir)
            if cnt:
                errors += cnt
                s += '"{:s}", '.format(k)

        if errors:
            print('{:d} errors encountered in {:s}'.format(errors, s[:-1]))
        else:
            print('All playlists successfully exported!')
        # print('dest '+repr(dest_dir)+' source '+repr(source_dir))

################################################################################

# noinspection PyPep8Naming
class pyWMP:
    """A class to create a more Python-y interface to Windows Media Player"""

    def __init__(self):
        # Get a pointer for Windows Media Player
        self.wmp = win32com_dispatch('WMPlayer.OCX', 0)

    def get_playlists(self, name=None):
        # Get playlists by name (default = All)
        if name:
            print('Getting songs for "' + name + '" playlist...'),
        else:
            print('Getting songs for all playlists...'),
        pl = pyWMPplaylist(self.wmp, name)
        print('Done!\n' + repr(pl))
        return pl

    def get_songs(self, playlist=None):
        # Get songs by playlist name (default = All)
        if playlist is None:
            print('Getting all songs...'),
            pl = self.wmp.mediaCollection.getByAttribute('MediaType', 'audio')
            pl = pyWMPsonglist([pl[i] for i in range(pl.count)], 'All Music')
        elif playlist:
            print('Getting songs for "' + playlist + '" playlist...'),
            pl = pyWMPplaylist(self.wmp, playlist)
            if pl:  # make sure playlist is valid
                pl = pl[playlist]
            else:
                pl = pyWMPsonglist([], 'playlist')
        else:
            raise NotImplemented("TODO: Find songs that are in no playlists")

        print('Done!\n' + repr(pl))
        return pl

    def remove_lists(self, min_songs=1):
        """List playlists and remove any that are empty (or less than the minimum number of songs specified)"""
        pl = pyWMPplaylist(self.wmp)

        empty = [(k, len(pl[k])) for k in pl.keys()]
        empty = [i for i in empty if i[1] < min_songs]

        print("Deleting playlists:"),

        for i in empty:
            if i[1] == 1:
                print("{0} ({1:d} song),".format(i[0], i[1])),
            else:
                print("{0} ({1:d} songs),".format(i[0], i[1])),

            del pl[i[0]]

    def list_broken(self, playlist=None, remove=False):
        # Check library for songs that don't exist

        if playlist:
            # search given playlist
            print('Searching playlist "' + playlist + '" for missing files...'),
            pl = pyWMPplaylist(self.wmp, playlist)
        else:  # search all songs in library
            print('Searching for missing files...'),
            pl = self.get_songs()

        bogus = []
        songs = self.get_songs(playlist)
        for i, s in enumerate(songs):
            if not os.access(s.sourceURL, os.F_OK):
                # remember song index if for bogus tracks
                bogus.append(i)

        if remove:
            # remove bogus songs in reverse order
            print('Removing', len(bogus), 'files:')
            bogus.reverse()
            if playlist:
                for i in bogus:
                    print(' ', songs[i].sourceURL)
                    pl[0].removeItem(pl[0][i])
            else:
                for i in bogus:
                    print(' ', songs[i].sourceURL)
                    self.wmp.mediaCollection.remove(songs[i], False)
        else:  # just show the results
            print('Done!')
            for i in bogus:
                print(' ', songs[i].sourceURL)


def test_wmp():
    """Simple function to test connection with Windows Media Player"""
    ptr = pyWMP()

    songs = ptr.get_songs()
    if len(songs) == 0:
        raise LookupError("No songs found!")

    pl = ptr.get_playlists()
    if len(pl) == 0:
        raise LookupError("No playlists found!")

    return ptr, songs, pl


################################################################################

if __name__ == '__main__':
    wmp, _, playlists = test_wmp()

    # Since simple file transfer doesn't work with newer Android devices, just export all the playlists
    # to manually copy over. Could support MTP file transfer in the future
    source_dir = os.path.join(r'C:\Users',os.environ['USERNAME'],'Music')
    if os.access(source_dir+os.path.sep+'Songs', os.F_OK):
        source_dir += os.path.sep+'Songs'
    # SD card has music and Playlists at the top level
    playlists.export_playlists(None, source_dir, r'..\music')
    os.popen('explorer '+os.curdir)

    def add_missing(root_dir='', extensions=('wav', 'wma', 'm4a', 'mp3')):
        # Ensure all songs on disk are in the library

        print('Gathering library information...'),
        songs = wmp.get_songs()
        path = {}
        for s in songs:
            if not os.access(s.sourceURL, os.F_OK):
                # ignore path if song doesn't exist
                continue
            else:  # add path
                tmp = s.sourceURL.split('\\')
                k = None
                for j in range(1, len(tmp)):
                    k = '\\'.join(s.sourceURL.split('\\')[:j])
                    if k not in path:
                        path[k] = []
                if k:
                    path[k].append(s.sourceURL)

        print(len(songs), 'songs in', len(path.keys()), 'directories.')
        del songs

        added = 0
        if root_dir:  # walk the directory to search for all songs
            print('Searching', root_dir, 'for missing songs (', '*.' + ', *.'.join(extensions), ')...')
            for root, dirs, files in os.walk(root_dir):
                if root in path:
                    for s in files:
                        f = root + '\\' + s
                        if s.split('.')[-1] not in extensions:
                            continue  # not a song
                        elif f in path[root]:
                            continue  # song accounted for
                        else:  # add song to library
                            print(' ', f.encode('utf8'))
                            added += 1
                            try:
                                wmp.wmp.mediaCollection.add(f)
                            except Exception as e:
                                print('broke on adding ', f)
                                print(e)
                else:  # whole directory missing
                    for s in files:
                        f = root + '\\' + s
                        if s.split('.')[-1] not in extensions:
                            continue  # not a song
                        else:  # add song to library
                            print(' ', f.encode('utf8'))
                            added += 1
                            wmp.wmp.mediaCollection.add(f)
        else:  # check each path for missing songs
            print('Searching directories for missing songs (', '*.' + ', *.'.join(extensions), ')...')
            for i in path.keys():
                songs = os.listdir(i)
                if len(songs) == len(path[i]):
                    continue  # all songs accounted for
                else:  # songs don't add up
                    for s in songs:
                        f = i + '\\' + s
                        if s.split('.')[-1] not in extensions:
                            continue  # not a song
                        elif f in path[i]:
                            continue  # song accounted for
                        else:  # add song to library
                            print(' ', f)
                            added += 1
                            wmp.wmp.mediaCollection.add(f)

        print('Done!', added, 'songs added')
