# perl-bluepages
A command-line interface to the IBM (internal) BluePages directory based on perl-ldap.

**If you do not have (a) access to the IBM internal network and (b) permission to use it, 
then you will not be able to use this program.**

Pre-requisites
--------------

- access to the IBM internal network
- perl 5.10 or later
- cpanm Bundle::Net::LDAP
- cpanm Clipboard
- cpanm Browser::Open

On macos you need the Command Line Developer Tools installed.  Calling "git" for the first time 
should automatically install them. 

Installation
------------

Try

    git clone https://github.com/thruston/perl-bluepages

then add something like this to your .bashrc

    alias pb="perl ~/perl-bluepages/pb.pl"

and create a local contacts directory

    mkdir ~/contacts

Adapt as appropriate for Windows.

Usage
-----

    pb --help

This program does command-line look up to bluepages with a local cache.
Each person found in bluepages is stored in a local VCF file (one each)
We look for people from the cache first and use them if we can.

Once you have found someone you can show their team, their reporting
chain, or add then to Contacts.app

Any successful look up (local or remote) copies the found person's Notes
email address to the clip board.

Toby Thurston -- 20 Nov 2016 
