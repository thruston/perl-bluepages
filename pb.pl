#! /usr/bin/perl -w
# Toby Thurston -- 20 Nov 2016 
# Command line interface to Bluepages

use strict;
use warnings;

# You need to install perl-ldap and Clipboard
use Net::LDAP::Entry;
use Net::LDAP;
use Clipboard;

# These are all standard with perl 5.10
use 5.010;
use LWP::Simple;
use Time::HiRes qw( gettimeofday tv_interval );
use Getopt::Long;
use Pod::Usage;
use Carp;
use MIME::Base64;
use Encode;

our $VERSION = '2.7'; 

=pod

=head1 NAME

pb.pl - Perl Bluepages: lookup people in the internal IBM directory.

This program will only work if you have access to the IBM internal network.

=head1 SYNOPSIS

    perl pb.pl usual name [country_code] [options]

=head1 ARGUMENTS and OPTIONS

In the example above "usual name" is the normal way of writing the person's name: 

    perl pb.pl toby thurston 

You can use UPPER or lower case.  If you are not sure how to spell the first name
just put the initial.  If you don't know the first name you can put "*", but you will
probably have to escape it with a backslash to stop your shell expanding it.

    perl pb.pl \* thurston

If you don't know the surname, then you can put the first few letters plus "*", but try not
to crash Blue Pages by asking for the entire world.  If the surname has several words
in it, you may need to use a "_" character instead of spaces in the surname.

    perl pb.pl Mark van_der_Pump

Instead of a name you can also try just an email address or a phone number.
Or you can use one of the following forms to search other fields:  

    serialnumber  085682866
    emailaddress  thurston\@uk.ibm.com
    tieline 430071
    telephonenumber 020-7021-9073
    mobile 07798-897287

Most of these can include * for matching in the value field. There is no field for
Mobex numbers in BP but most people put them in the mobile field in (brackets) after
the main mobile number. Creative use of *s may help.

"country code" is "gb" or "us" or "fr" or "de" etc  
Note that country code "uk" is for Ukraine; use "gb" for United Kingdom.

=over 4 
   
=item --force

go directly to BP instead of showing any local cached version

=item --quiet

turns off any warning or information messages

=item --chain

shows the list of people to whom the person found reports

=item --peers

shows the list of people with the same manager

=item --global

modifies --chain and --peers to use "Global Team Leader" instead of "Manager"

=item --team

shows the list of people reporting to the person found (if any)

=item --asst

shows the full details for the person's assistant (if any)

=item --notes

shows just Notes addresses (and also in any chain or team)

=item --email

shows just Internet email addresses (& in any chain or team)

=item --svpic

saves a copy of the person's bluepages picture alongside the vcf (as well as putting it *in* the vcf)

=item --nopic

do not get the person's bluepages picture at all

=item --show "format" 

for use in scripts, get a VCF card and show fields (implies quiet)

=item --bluepages

show person in Bluepages

=item --link

put URL for person in Bluepages on clip board

=item --delete

remove local VCF file (with confirmation unless --quiet)

=item --usage, --help, --man

Show increasing amounts of help text, and exit.

=item --version

Print version and exit.

=back

=head1 DESCRIPTION

Any successful look up will also put the person's Notes address on the clip board. Although if you
have said "--email" it will be the external email address instead, and if you say "--card" it will
be a neat text representation of a business card.

You can only have one of --chain, --team, or --peers at the same time.
Similarly it's either --notes or --email.
If you specify --global without one of --chain or --peers, you just get an indication if 
there is a global team lead defined or not.

--show implies --quiet and cancels --phone, --chain, --team, and --peers

=head1 AUTHOR

Toby Thurston -- 20 Nov 2016 

=cut

# Process the option switches
my $Show_peers              = 0;
my $Show_team               = 0;
my $Show_assistant          = 0;
my $Show_chain              = 0; my %Shown_in_chain=(); my $Chain_indent=0;
my $Want_email              = 0;
my $Want_notes              = 0;
my $Want_card_on_clip       = 0;
my $Want_short_card_on_clip = 0;
my $Want_BP_link_on_clip    = 0;
my $Add_to_clip             = 0;
my $Save_picture            = 0;
my $Skip_picture            = 0;
my $Keep_quiet              = 0;
my $Search_BP               = 0;
my $Search_locally          = 1;
my $Show_format             = '';
my $Show_BP                 = 0;
my $Kill_VCF                = 0;
my $Dump_raw_data           = 0;
my $World_wide_search       = 0;
my $Use_gtl_for_manager     = 0;

my $options_ok = GetOptions(
    'show:s'  => \$Show_format,
    add       => \$Add_to_clip,
    asst      => \$Show_assistant,
    bluepages => \$Show_BP,
    card      => \$Want_card_on_clip,
    chain     => \$Show_chain,
    delete    => \$Kill_VCF,
    dump      => \$Dump_raw_data,
    email     => \$Want_email,
    force     => \$Search_BP,
    global    => \$Use_gtl_for_manager,
    link      => \$Want_BP_link_on_clip,
    nopic     => \$Skip_picture,
    notes     => \$Want_notes,
    peers     => \$Show_peers,
    quiet     => \$Keep_quiet,
    scard     => \$Want_short_card_on_clip,
    svpic     => \$Save_picture,
    team      => \$Show_team,
    tree      => \$Show_team,
    world     => \$World_wide_search,
    
    'version'     => sub { warn "$0, version: $VERSION\n"; exit 0; }, 
    'usage'       => sub { pod2usage(-verbose => 0, -exitstatus => 0) },                         
    'help'        => sub { pod2usage(-verbose => 1, -exitstatus => 0) },                         
    'man'         => sub { pod2usage(-verbose => 2, -exitstatus => 0) },

) or die pod2usage();
die pod2usage unless @ARGV;

$Show_team      = 0 if $Show_chain;
$Show_team      = 0 if $Show_peers;
$Show_chain     = 0 if $Show_peers;
$Want_email     = 0 if $Want_notes;
$Show_peers     = 0 if $Show_format;
$Show_team      = 0 if $Show_format;
$Show_assistant = 0 if $Show_format;
$Show_chain     = 0 if $Show_format;
$Keep_quiet     = 1 if $Show_format;
$Search_locally = 0 if $Search_BP;
$Search_locally = 0 if $Save_picture;
$Kill_VCF       = 0 unless $Search_locally;

# Some Globals
my $Contacts_dir = get_home_path() . '/contacts/' ;
my %VCF_key_for  = (
    serial  => 'X-CNUM',
    div     => 'X-DIV',
    dept    => 'X-ODB',
    asst    => 'X-ASST',
    asfilter=> 'X-ASFILTER',
    country => 'X-COUNTRY',
    manager => 'X-MANAGER',
    glTeamLead => 'X-GLTEAMLEAD',
    manflag => 'X-ISMANAGER',
    notes   => 'EMAIL;NOTES',
    telwk   => 'TEL;WORK',
    telmb   => 'TEL;MOBILE',
    telti   => 'TEL;!TIE',
    telmx   => 'TEL;MOBEX',
    vcfname => '_SKIP_',
    );

my @Useful_BP_Fields = qw(
    c
    callupname
    dept
    div
    emailaddress
    givenname
    glTeamLead
    internalmaildrop
    ismanager
    jobresponsibilities
    locationcity
    manager
    mobile
    notesemail
    secretary
    serialnumber
    sn
    telephonenumber
    tieline
);

my %Full_Field_Name_for = (
    job   => 'jobresponsibilities',
    tie   => 'tieline',
    phone => 'telephonenumber',
    tel   => 'telephonenumber',
    mob   => 'mobile',
    location => 'locationcity',
);

my %Correction_of = (  # keep keys all lowercase 
    rochard => 'Richard',

);

my %is_a_BP_field = map { $_ => 1 } (@Useful_BP_Fields, keys %Full_Field_Name_for);

my $IBM_email_pattern = qr{ \A [-_.A-Za-z0-9]+\@([a-z][a-z])1?\.ibm\.com \Z }ix;  # 1? allows for hk1
my $IBM_notes_pattern = qr{ \A ([A-Za-z].*?)\/(.*?)\/(IBM|IDE)         \Z }ix;
my $IBM_notes_long_pat= qr{ \A (\S.*?\/IBM)\@([A-Z]{5})  \Z }ix;
my $canonical_pattern = qr{ \A CN= }ix;

# Read the args (and play about with them)
#
# We want to accept the following forms of input
# - firstname surname [country]
# - vcfname
# - email address
# - notes address
# - phone number
# - bpfield_name field_value

my %desiderata = ( );

if ( -f "@ARGV" ) {
    my $file = "@ARGV"; @ARGV=();
    warn "Refreshing --> $file\n" unless $Keep_quiet;

    my $p = load_person_from_vcf($file);
    $desiderata{emailaddress} = $p->{email};

    $Search_BP = 1;
    $Search_locally = 0;
    $Keep_quiet = 1;
    $Show_format = 'NOTHING';
}

elsif ( "@ARGV" =~ m{$IBM_email_pattern} ) { $desiderata{emailaddress} = splice @ARGV; $desiderata{c} = $1 eq 'uk' ? 'gb' :$1 }
elsif ( "@ARGV" =~ m{$IBM_notes_pattern} ) { $desiderata{notesemail}   = join ' ', splice @ARGV }
elsif ( "@ARGV" =~ m{$IBM_notes_long_pat}) { 
    $desiderata{notesemail} = $1;
    my $domain = $2;
    if ( $domain eq 'LOTUS' ) {
        $desiderata{c} = 'us';
    } 
    elsif ( $domain =~ m{\AIBM([A-Z]{2})\Z}i ) {
        $desiderata{c} = $1;
    }
    @ARGV = ();
}
elsif ( "@ARGV" =~ m{$canonical_pattern} ) { $desiderata{notesemail}   = join ' ', splice @ARGV }
elsif ( "@ARGV" =~ m{\A\d+\Z}            ) { $desiderata{tieline}      = splice @ARGV }
else {
    while ( @ARGV ) {
        last unless $is_a_BP_field{$ARGV[0]};
        my $k = shift;
        if ( $Full_Field_Name_for{$k} ) { $k = $Full_Field_Name_for{$k} }
        $desiderata{$k} = shift;
    }
}

if ( @ARGV ) {
    my $first_name   = shift;
    my $surname      = shift || '*' ;
    my $country_code = shift;

    if ( exists $Correction_of{$first_name} ) {
        $first_name = $Correction_of{lc($first_name)}
    }

    if ( $surname =~ /\A(van|de|von|le)\Z/i && defined $country_code && length($country_code) > 0 ) {
        $surname = "$surname $country_code";
        $country_code = shift;
        if ( $surname =~ /\A(van der|de la)\Z/i && defined $country_code && length($country_code) > 0 ) {
            $surname = "$surname $country_code";
            $country_code = shift;
        }
    }
    $surname =~ s/_/ /g; # van_der_pump
    if ( !defined $country_code) {
        $country_code = get_cc();
    } 
    elsif ( -f $country_code ) { # because you passed a * which has matched a file...
        $country_code = '*';
    }	
    elsif ( $country_code eq '..' ) {
        $country_code = '*'
    }
    $desiderata{c}         = $country_code;
    $desiderata{sn}        = $surname;
    $desiderata{givenname} = $first_name;
}

# Find the person
my $p; # person (to be) found
my $need_to_save_vcf;
my $Local_vcf_name;
my $search_start_time = [ gettimeofday ];

if ( $Search_locally ) {
    $p = find_person_in_local_files(\%desiderata);
}

if ( ! defined $p ) {
    # add local cc to reduce stress on BP
    $desiderata{c} = get_cc() unless defined $desiderata{c};
    # make a filter
    my $f = '(&';
    for my $k (sort keys %desiderata ) {
        my $v = $desiderata{$k};
        if ($k eq 'notesemail') {
            $v = canonize_notes($v);
        }
        $f .= "($k=$v)";
    }
    $f .= ')';
    warn "Looking for $f\n" unless $Keep_quiet;

    if ( $Dump_raw_data ) {
        show_all_data_for_filter_in_bluepages($f);
        exit;
    }


    # search BP with it
    $p = find_person_in_bluepages($f);
    $need_to_save_vcf++;
}

if ( $Show_assistant ) {
    die "No assistant for $p\n" unless $p->{asfilter};
    $p = find_person_in_bluepages($p->{asfilter});
    $need_to_save_vcf++;
}

# show how long it took
warn sprintf "Time: %d ms\n", int(0.5+1000*tv_interval($search_start_time)) unless $Keep_quiet;

# Display some results
print get_person_lines($p) unless $Show_format;

if ( $Show_chain ) {
    die "No manager defined\n" unless $p->{manager};

    my $ldap = Net::LDAP->new('bluepages.ibm.com') or die "Not on IBM network ---> $@\n";
    $ldap->bind;
    show_chain($ldap, $p->{manager});
    $ldap->unbind;
}

if ( $Show_peers ) {
    die "No manager defined\n" unless $p->{manager};

    my $ldap = Net::LDAP->new('bluepages.ibm.com') or die "Not on IBM network ---> $@\n";
    $ldap->bind;
    show_team($ldap, $p->{manager});
    $ldap->unbind;
}

if ( $Show_team ) {
    die "Not a manager\n" unless $p->{manflag};

    my $ldap = Net::LDAP->new('bluepages.ibm.com') or die "Not on IBM network ---> $@\n";
    $ldap->bind;
    show_team($ldap, $p);
    $ldap->unbind;
}

# Save the VCF if we need to
$Local_vcf_name = save_person_to_vcf($p) if $need_to_save_vcf;

if ( $Kill_VCF ) {
    warn "=======> Delete $Local_vcf_name? Are you sure? (y/N) <=======\n";
    my $response = <STDIN>;
    unlink $Local_vcf_name if ($response =~ /\Ay/io);
    exit;
}

# Return a one liner if $Show format is set
if ( $Show_format ) {
    $Show_format =~ s[\%f][$Local_vcf_name]eig;
    $Show_format =~ s[\%n][$p->{fn}]eig;
    $Show_format =~ s[\%c][$p->{serial}]eig;
    $Show_format =~ s[\%e][$p->{email}]eig;
    $Show_format =~ s[\%s][$p->{notes}]eig;
    $Show_format =~ s[\%j][$p->{title}]eig;
    $Show_format =~ s[\%tt][$p->{telti}]eig;
    $Show_format =~ s[\%tx][$p->{telmx}]eig;
    $Show_format =~ s[\%tm][$p->{telmb}]eig;
    $Show_format =~ s[\%tw][$p->{telwk}]eig;
    print "$Show_format\n" unless uc($Show_format) eq 'NOTHING';
    Clipboard->copy($Show_format);
} 
else {
    print "\n";
}

# Show in BP if wanted
my $url = 'http://w3.ibm.com/bluepages/searchByName.wss?uid=' . $p->{serial} . '&task=viewrecord';
if ( $Show_BP ) {
    my $command = $^O eq 'MSWin32' ? 'start "Blue Pages"' : 'open';
    system $command . ' "' . $url . '"';
}
if ( $Want_BP_link_on_clip ) {
    Clipboard->copy($url);
}

exit;
#****************************************************************************
#* End of main line                                                         *
#****************************************************************************

sub find_person_in_local_files {

    my $needle;
    my @local_vcfs;
    my $desired_ref = shift;
    my @bp_keys = sort keys %$desired_ref;
    if ( "@bp_keys" eq 'c givenname sn' ) {
        $desired_ref->{givenname} .= '*' if length($desired_ref->{givenname}) == 1
                                         && $desired_ref->{givenname} ne '*';
        @local_vcfs = glob($Contacts_dir . make_vcf_name(@$desired_ref{qw(sn givenname c)}));
    }
    elsif ( "@bp_keys" eq 'c emailaddress' ) {
        $needle = $desired_ref->{emailaddress};
        warn "Grepping local files for $needle . . . \n" unless $Keep_quiet;
        @local_vcfs = `grep -il "$needle" ${Contacts_dir}*.vcf`;
    }
    else {
        ($needle) = values %$desired_ref;  # this is likely to produce rubbish, first of random order of values...
        warn "Grepping local files for $needle . . . \n" unless $Keep_quiet;
        @local_vcfs = `grep -il "$needle" ${Contacts_dir}*.vcf`;
    }

    return unless @local_vcfs;

    chomp(@local_vcfs);
    my $choice = 0;
    if ( @local_vcfs > 1) {
        my $i = 0;
        for (@local_vcfs) {
            printf "%2d. %s\n", $i++, $_;
        }
        print "Please chose one of these people.\n";
        $choice = <STDIN>;
        return unless $choice =~ /\A\d+\Z/ && $choice < $i;
    }

    return unless -e $local_vcfs[$choice];

    warn "Using local copy -> $local_vcfs[$choice]\n" unless $Keep_quiet;
    $Local_vcf_name = $local_vcfs[$choice];
    return load_person_from_vcf($Local_vcf_name);
}


sub show_all_data_for_filter_in_bluepages {
    my $ldap = Net::LDAP->new('bluepages.ibm.com') or die "Not on IBM network ---> $@\n";
    $ldap->bind ;

    my $filter = shift;
    my $search = $ldap->search(base   => "ou=bluepages,o=ibm.com",
                              filter => $filter );

    $search->code  && die $search->error, "\n";

    die "No-one found with $filter\n" unless $search->count;
    die "Ambiguous filter\n"          unless 1==$search->count;

    use Data::Dumper;
    print Dumper($search);

    $ldap->unbind;

}
sub find_person_in_bluepages {

    # Now we have to actually go to Blue Pages...
    my $ldap = Net::LDAP->new('bluepages.ibm.com') or die "Not on IBM network ---> $@\n";
    $ldap->bind ;

    my $filter = shift;
    my $search = $ldap->search(base   => "ou=bluepages,o=ibm.com",
                              attrs  => \@Useful_BP_Fields,
                              filter => $filter );

    $search->code  && die $search->error, "\n";

    die "No-one found with $filter\n" unless $search->count;

    my @entries = $search->entries;
    my $choice = 0;
    my %short_id_for = ();
    if ( @entries > 1 ) {
        for (@entries) {
            $short_id_for{$_} = get_short_id_from_ldap_entry($_);
        }
        @entries = map  { $_->[1] }
                   sort { uc($a->[0]) cmp uc($b->[0]) }
                   map  { [ $short_id_for{$_}, $_ ] } @entries;
        my $i = 0;
        for (@entries) {
            printf "%2d. %s\n", $i++, $short_id_for{$_};
        }
        print "Please chose one of these people.\n";
        $choice = <STDIN>;
        die "Quit\n" unless $choice =~ /\A\d+\Z/ && $choice < $i;
    }

    my $e = $entries[$choice];
    # attempt to extract a mobex number out of the mobile field
    my ($mobile,$mobex);
    $mobile = gkv($e,'mobile');
    # warn "Raw mobile field: $mobile\n" unless $Keep_quiet;

    if ( $mobile =~ /\A(.+?\S)\s*\(.*?(2[67]\-?\d\d\d\d)\s*\)\Z/io ) {
        $mobile = $1;
        $mobile =~ s/mobex//io;
        $mobex  = $2;
        $mobex  =~ s/-//o;
    }
    elsif ( $mobile =~ /\A(.+?)MO?B?E?X.*?(2[67]\d\d\d\d)/io ) {
        $mobile = $1;
        $mobex = $2;
    }
    elsif ( $mobile =~ /7967\-?(27\d\d\d\d)/io ) {
        $mobex  = $1;
    }
    elsif ( $mobile =~ /\A(.+?\S)\s*\(\s*(\d\d\d\d\d\d)\s*\)\Z/io ) {
        $mobile = $1;
        $mobile =~ s/mobex//io;
        $mobex  = $2;
    }
    # 44-(0)7725-829925,  x37278984
    elsif ( $mobile =~ /\A(.+?\S)\s+.*?(3727\d\d\d\d)/io ) {
            $mobile = $1;
            $mobex = $2;
    }
    else {
      warn "No MOBEX in mobile number >>>>>>>>> $mobile\n" if $mobile && !$Keep_quiet;
    }

    if ( $mobex && $mobex !~ /\A37/) {
        $mobex = '37'.$mobex
    }


    my @names = get_names_from_ldap_entry($e);
    my $n       = join ';', @names;
    my $fn      = $names[1].' '.$names[0];
    my $serial  = gkv($e,'serialnumber');
    my $div     = gkv($e,'div');
    my $dept    = gkv($e,'dept');
    my $title   = gkv($e,'jobresponsibilities');
    my $email   = gkv($e,'emailaddress');
    my $notes   = format_notes_address(gkv($e,'notesemail'));
    my $telwk   = format_phone_number(gkv($e,'telephonenumber'));
    my $telmb   = format_phone_number($mobile);
    my $telti   = format_ibm_phone_number(gkv($e,'tieline'));
    my $telmx   = format_ibm_phone_number($mobex);
    my $adr     = gkv($e,'internalmaildrop').' '.gkv($e,'locationcity');
    my $vcfname = make_vcf_name($names[0], $names[1], gkv($e,'c'));
    my $manager = gkv($e,'manager');
    my $manflag = 'Y' eq gkv($e,'ismanager')?1:0;
    my $country = gkv($e,'c');
    my $asfilter= gkv($e,'secretary');
    my $asst    = 'None';

    if ( $asfilter ) {
        $asfilter =~ s/,ou=bluepages,o=ibm.com\Z//;
        $asfilter = attrib_list_to_ldap_filter($asfilter);
        $search = $ldap->search(base   => 'ou=bluepages,o=ibm.com',
                                filter => $asfilter,
                                attrs  => ['tieline', 'notesemail' ] );
        $search->code  && die $search->error, "\n";
        if ( $search->count == 0 ) {
            $asst = $asfilter . ' <--- Not found';
        } else {
            my $e = $search->entry(0);
            $asst = format_notes_address(gkv($e,'notesemail')).' ('.format_ibm_phone_number(gkv($e,'tieline')).')';
        }
    }

    $ldap->unbind;   # take down session

    if ( $telwk eq $telmb ) {
        $telwk = ''
    }

    return {
        n       => $n        ,
        fn      => $fn       ,
        serial  => $serial   ,
        dept    => $dept     ,
        div     => $div      ,
        title   => $title    ,
        email   => $email    ,
        notes   => $notes    ,
        telwk   => $telwk    ,
        telmb   => $telmb    ,
        telti   => $telti    ,
        telmx   => $telmx    ,
        adr     => $adr      ,
        vcfname => $vcfname  ,
        manager => $manager  ,
        manflag => $manflag  ,
        country => $country  ,
        asst    => $asst     ,
        asfilter=> $asfilter ,
    };
}

sub format_notes_address {
    my $s = shift; return '' unless $s;
    $s =~ s/O=//;
    $s =~ s/OU=//g;
    $s =~ s/CN=//;
    $s =~ s/\@IBM..\Z//;
    return $s;
}

sub format_ibm_phone_number {
    my $s = shift; return '' unless $s;
    $s =~ s/\A7-//;
    return $s;
}
sub format_phone_number {
    my $s = shift; return '' unless $s;

    $s =~ s/-//g;
    $s =~ s/\s+//g;
    $s =~ s/\(//g;
    $s =~ s/\)//g;
    $s =~ s/,//g;

    if ( $s =~ m{\A (0|44)}x  ) { # UK

        $s =~ s/^440/44/;
        $s =~ s/^0/44/;

        return $s unless length($s)>11;

        my $n = substr($s,2,1);
        if    ( $n == 1 ) { $s = '+44-1'.substr($s,3,3).'-'.substr($s,6,3).'-'.substr($s,9,3).' '.substr($s,12) }
        elsif ( $n == 2 ) { $s = '+44-2'.substr($s,3,1).'-'.substr($s,4,4).'-'.substr($s,8,4).' '.substr($s,12) }
        elsif ( $n == 7 ) { $s = '+44-7'.substr($s,3,3).'-'.substr($s,6,3).'-'.substr($s,9,3).' '.substr($s,12) }
        else { $s = '+'.$s }
    }
    elsif ( $s =~ m{\A 33}x ) { # France
       if ( $s =~ m{\A330?(\d)(\d\d)(\d\d)(\d\d)(\d\d)\Z} ) {
           $s = sprintf "33-%s-%s-%s-%s-%s", $1, $2, $3, $4, $5;
       }
       $s = '+'.$s
    }
    elsif ( $s =~ m{\A 1}x ) { #us etc
       if ( $s =~ m{\A1(\d\d\d)(\d\d\d)(\d\d\d\d)\Z} ) {
           $s = sprintf "1-%s-%s-%s", $1, $2, $3;
       }
       $s = '+'.$s

    }
    else {
        $s = '+'.$s;
    }

    return $s;
}

sub save_person_to_vcf {
    my $p = shift;
    my $local_vcf_name = $Contacts_dir . $p->{vcfname};
    my @vcf_contents   = create_vcf_lines_from($p);

    open my $vcf_handle, '>', "$local_vcf_name" or die "Can't save VCF to file $local_vcf_name. $!\n";
    print   $vcf_handle @vcf_contents;
    close   $vcf_handle;

    return $local_vcf_name;
}

sub create_vcf_lines_from {
    my $p = shift;
    my @out = ();
    push @out, "BEGIN:vCard\nVERSION:3.0\nORG:IBM\n";

    for (sort keys %$p ) {
        my $key = $VCF_key_for{$_} || uc($_);
        next if $key eq '_SKIP_';
        push @out, $key.':'.$p->{$_}."\n";
    }

    if ( ! $Skip_picture ) {
        my $serial = $p->{serial};
        my $pic = get("http://w3.ibm.com/bluepages/photo/ImageServlet.wss/$serial.jpg?cnum=$serial");
        if ( defined $pic && substr($pic,0,5) ne 'GIF89' ) {  # don't store "Missing.gif"

            $pic =~ s/\x00*\Z//; # no trailing nulls thank you Bluepages

            if ( $Save_picture ) {
                open my $jpg_handle, '>', "$serial.jpg";
                binmode $jpg_handle;
                print $jpg_handle $pic;
                close $jpg_handle;
            }

            $pic = 'PHOTO;ENCODING=BASE64;TYPE=JPEG:'.encode_base64($pic,'');
            while ( length($pic) > 76 ) {
                push @out, substr($pic,0,76), "\n  "; # <--- NB spaces at start of next line...
                $pic = substr($pic,76);
            }
            push @out, "$pic\n" if $pic;
        }
    }

    my ($d, $m, $y) = (localtime)[3..5];
    push @out, sprintf "REV:%04d-%02d-%02d\n", 1900+$y, 1+$m, $d;
    push @out, "END:vCard\n";
    return @out;
}

sub load_person_from_vcf {
    my $local_vcf_name = shift;
    my %p = ();
    my %names = (); while ( my ($k, $v) = each %VCF_key_for ) { $names{$v} = $k }

    open my $vcf_handle, '<', $local_vcf_name;
    while (<$vcf_handle>) {
        my ($k, $data) = /^([-A-Z\!\;\=\"]+)\:\s*(.*?)\s*$/;
        next unless $k; # skip lines with no key: at the start (continuation lines...)
        $k = $names{$k} || lc($k);
        $p{$k} = $data;
    }
    close $vcf_handle;
    return \%p;
}

sub make_vcf_name {
    my $s = join '-', @_;
    $s =~ s/\s+/_/g; # remove any spaces in the name
    return $s . '.vcf';
}

sub show_chain {
    my $ldap = shift;
    my $q = shift;
    return if $Shown_in_chain{$q}++;
    $Chain_indent++;

    my $search = $ldap->search (  base   => '', filter => attrib_list_to_ldap_filter($q) );

    $search->code    && die $search->error;
    $search->entries || die "No manager found with filter based on $q\n";

    my $e = $search->entry(0);

    my $address = $Want_email ? gkv($e,'emailaddress') : format_notes_address(gkv($e,'notesemail')); 

    if ( $Want_notes || $Want_email ) {
        my $s = Clipboard->paste;
        Clipboard->copy("$s, $address"); # append managers email address to Clip
        print ",\n$address";
    } else {
        print "\n";                                      # allow for Van der Pump...
        my ($sn, $cn) = gkv($e,'callupname') =~ /^\s*(\S.+)\,\s*(\S.*)/;

        print '  ' x $Chain_indent, $cn, ' ', $sn, ', ',gkv($e,'jobresponsibilities'), "\n";
        print '  ' x $Chain_indent, 'O:', gkv($e,'div'), '/', gkv($e,'dept'), ' E:', $address, "\n";
     }

    show_chain($ldap, gkv($e,'manager'));
    return;
}

sub show_team {
    my $ldap = shift;
    my $p = shift;

    if ( ref $p ne 'HASH' ) { # if $p is not a hash ref then it should be a BP attrib list...
        $p = find_person_in_bluepages(attrib_list_to_ldap_filter($p));
    }

    my $cc = $p->{country};
    my ($serial, $country) = $p->{serial} =~ /^(\S{6})(\d\d\d)$/;

    my $query = $ldap->search (  base   => "ou=bluepages,o=ibm.com",
                                 attrs  => [ "c", "serialnumber", "ismanager", "callupname", "div",
                                             "dept", "emailaddress", "notesemail", "jobresponsibilities" ],
                                 filter => "(&(c=$cc)(managercountrycode=$country)(managerserialnumber=$serial))" );

    $query->code && die $query->error;
    $query->entries || die "No employees found for $serial in $country\n";

    my $r = $query->as_struct;
    for (sort { $r->{$a}->{callupname}[0] cmp $r->{$b}->{callupname}[0] } keys %{$r} ) {

        my $e = $r->{$_};

        my $address = $Want_email ? $e->{emailaddress}->[0]
                                  : format_notes_address($e->{notesemail}->[0]) ;

        if ( $Want_notes || $Want_email ) {
            next unless $address;
            my $s = Clipboard->paste;
            Clipboard->copy("$s, $address"); # append managers email address to Clip
            print ",\n$address";
        } else {
            my $who = $address || $e->{callupname}->[0];
            print '  ',
            $e->{div}->[0] || '?' ,'/',
            $e->{dept}->[0],'|',
            $who,
            $e->{jobresponsibilities} ? ', ('.$e->{jobresponsibilities}->[0].')' : '',
            $e->{ismanager}->[0] eq 'Y' ? ' *' : '',
            "\n";
        }
    }
    return;
}

sub get_person_lines {
    my $p = shift;

    my $sep = $Want_card_on_clip || $Want_short_card_on_clip ? "\n\n" : ', ';
    my $clip = $Add_to_clip ? Clipboard->paste . $sep : '';

    Clipboard->copy($clip.$p->{notes}); # put the Notes name on the clipboard as default

    if ( $Want_notes ) {
        return $p->{notes}
    }

    if ( $Want_email) {
        Clipboard->copy($clip.$p->{email});
        return $p->{email}
    }

    my @out = ();
    push @out, $p->{fn};
    push @out, ", $p->{title}" if $p->{title};
    push @out, "\n";
    push @out, "S: $p->{serial} O: $p->{div}/$p->{dept}\n" if $p->{serial} && $p->{dept} && $p->{div};
    push @out, "E: $p->{email}\n" if $p->{email};
    push @out, "N: $p->{notes}\n" if $p->{notes};

    if ($p->{telwk}) {
        my $out = "T: $p->{telwk}"; 
        $out .= " ($p->{telti})" if $p->{telti};
        push @out, "$out\n";
    }

    if ($p->{telmb}) {
        my $out = "T: $p->{telmb}"; 
        $out .= " ($p->{telmx})" if $p->{telmx};
        push @out, "$out\n";
    }

    push @out, "MP: $p->{adr}\n"  if $p->{adr};
    push @out, "Assistant: $p->{asst}\n" if $p->{asst};

    Clipboard->copy($clip."@out") if $Want_card_on_clip;

    if ( $Want_short_card_on_clip ) {
        Clipboard->copy($clip."$p->{fn}, $p->{title}\n$p->{email}\n$p->{telmb}\n");
    }

    return @out;
}

sub attrib_list_to_ldap_filter {
    my $attribute_list = shift;
    my $filter = '';
    $filter .= "($_)" for split ",", $attribute_list;
    return '(&'.$filter.')';
}

sub get_short_id_from_ldap_entry {
    # in <-- a Net::LDAP::Entry object
    # out -> a short string identifying this entry
    my $e = shift;
    return '' unless $e && ref $e;
    my $n = get_names_from_ldap_entry($e);
    my $job = $e->get_value('jobresponsibilities') ? ' - '.gkv($e,'jobresponsibilities') : '';
    my $ser = gkv($e,'serialnumber','?');
    my $odb = gkv($e,'dept','?');
    my $nnn = $Want_email         ? gkv($e,'emailaddress')
            : $desiderata{mobile} ? format_phone_number(gkv($e,'mobile'))
            :                       format_notes_address(gkv($e,'notesemail'));
    return sprintf "%-64s S:%s ODB:%-6s %s", substr($n.$job,0,64), $ser, $odb, $nnn;
}

sub get_names_from_ldap_entry {
    # There is a long and sorry history to this.  sn is OK in BP as surname, but the given name
    # field is variously abused so our current punt is that the 1st word of the notes email is
    # the best source of the first name.
    return unless defined wantarray;

    my $e = shift;
    return unless $e && ref $e;

    my $first_name;
    if ( gkv($e,'notesemail') =~ /\ACN=([A-Z][a-z]+)/o ) {
        $first_name = $1;
    } elsif ( gkv($e,'callupname') =~ /\(([A-Z][a-z]+)\)/o ) {
        $first_name = $1;
    } else {
        my @names = grep {!/\./} $e->get_value('givenname'); # no initials thanks
        $first_name = $names[0] || gkv($e,'givenname'); # unless that's all there is
    }
    my $last_name = gkv($e,'sn');
    return ( $last_name, $first_name ) if wantarray;
    return "$first_name $last_name";
}

sub gkv {
    my $e = shift; confess unless $e && ref $e;
    my $k = shift;
    my $default = shift || '';
    my $value = $e->get_value($k) || $default;
    return decode('utf8',$value);
}

sub canonize_notes {
    my $s = shift;
    return qq{*$s*} if $s =~ m{\A CN= }x;

    $s =~ s/\//\/OU=/g;
    $s =~ s/OU=IBM/O=IBM/;
    return qq{*$s*};

}

sub get_cc {
    return '*' if $World_wide_search;
    
    if ( defined $desiderata{emailaddress} ) {
        my ($cc) = $desiderata{emailaddress} =~ m{ \@([a-z]{2}).ibm.com$ }x;
        $cc = 'gb' if $cc eq 'uk';
        return $cc
    }

    if ( defined $ENV{LANG} && $ENV{LANG} =~ m{ \A [A-Z][A-Z]\_([A-Z][A-Z]) }ix) {
        return lc($1);
    }
    
    return 'gb'
}

sub get_home_path {
    if ( defined $ENV{HOME}      ) { return $ENV{HOME}      }
    if ( defined $ENV{HOMEDRIVE} ) { return $ENV{HOMEDRIVE} }
    return 'c:'
}
