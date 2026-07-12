#!/usr/bin/perl
use strict;
use IO::Socket::INET;

my $port = $ARGV[0] || 3000;
my %mime = (
  html => 'text/html', css  => 'text/css',  js   => 'application/javascript',
  jpg  => 'image/jpeg', jpeg => 'image/jpeg', png  => 'image/png',
  gif  => 'image/gif',  svg  => 'image/svg+xml', ico => 'image/x-icon',
  woff => 'font/woff',  woff2 => 'font/woff2',  ttf => 'font/ttf',
);

my $server = IO::Socket::INET->new(
  LocalPort => $port, Listen => 10, ReuseAddr => 1, Proto => 'tcp'
) or die "Cannot bind to port $port: $!\n";

print "Serving on http://localhost:$port\n";

while (my $client = $server->accept) {
  my $request = '';
  while (my $line = <$client>) {
    $request .= $line;
    last if $line eq "\r\n";
  }

  my ($path) = $request =~ /^GET\s+(\S+)/;
  $path //= '/';
  $path =~ s/\?.*//;
  $path =~ s|^/||;
  $path = 'index.html' if $path eq '' || $path =~ m|/$|;

  if (-f $path) {
    my ($ext) = $path =~ /\.(\w+)$/;
    my $type = $mime{lc($ext // '')} || 'application/octet-stream';
    open my $fh, '<:raw', $path or do {
      print $client "HTTP/1.0 403 Forbidden\r\n\r\n"; close $client; next;
    };
    local $/; my $body = <$fh>; close $fh;
    print $client "HTTP/1.0 200 OK\r\nContent-Type: $type\r\nContent-Length: " . length($body) . "\r\n\r\n$body";
  } else {
    print $client "HTTP/1.0 404 Not Found\r\nContent-Type: text/plain\r\n\r\nNot found: $path";
  }
  close $client;
}
