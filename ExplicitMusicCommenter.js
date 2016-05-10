/* 	Rename me to Explicit.js
	Double Click in Explorer to run
Down below replace
badword1|badword2|badword3
with your own badwords add more with a | between the words
Script by Joseph Liszovics for WPOB-FM: http://wpob.com       */

var iTunesApp = WScript.CreateObject("iTunes.Application");
var tracks    = iTunesApp.LibraryPlaylist.Tracks;
var numTracks = tracks.Count;
var i;
for (i = 1; i <= numTracks; i++)
{
	try {
		//get the track and lyrics from the track
		var currTrack = tracks.Item(i);
		var mylyrics = currTrack.Lyrics;

		// look for bad words
		var badFound = mylyrics.match(/nigger|nigga|fuck|cocksucker|mother-fucker|fuck|shit|cunt|cock|pussy|dick|weed|bitch|dike|faggot/i);

		// flag the track as Explicit
		if (badFound != null) {
			var myComm = currTrack.Comment;
			if (myComm.match(/explicit/i) == null)
				currTrack.Comment = myComm + " Explicit";
		}
	} catch(er) {
	}
}