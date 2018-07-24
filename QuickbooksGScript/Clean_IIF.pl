# Run this script to clean up the IIF file created by the Google Script.

# The file downloaded from Google Drive has the first three characters as \357\273\277, denoting something about it
# being a UTF-8 file. This script strips those off and puts the tab characters in so I no longer have to do this manually
# in vim. Woot!
while (<>) {
  #chomp($_);
  $_ =~ s/\357\273\277//;
  $_ =~ s/XXXXXX/\t/g;
  print $_;
  }
print "\r\n";