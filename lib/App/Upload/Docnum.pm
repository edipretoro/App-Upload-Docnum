package App::Upload::Docnum;

use strict;
use warnings;

use Spreadsheet::Read;
use Getopt::Long::Descriptive;

use Progress::Any;
use Progress::Any::Output;
Progress::Any::Output->set('TermProgressBarColor');

use DBI;
use File::Copy;

use Data::Printer;

__PACKAGE__->run() unless caller();

sub run {
  my ( $options, $usage ) = describe_options(
    '%c %o',
    [ 'spreadsheet|s=s', 'path to the spreadsheet', { required => 1 } ],
    [ 'directory|d=s', 'path to the directory where the PDF are stored', { required => 1 } ],
    [ 'output|o=s', 'path where PMB store the PDF files', { required => 1 } ],
    [ 'dsn|c=s', 'DSN to connect to the database', { required => 1 } ],
    [ 'user|u=s', 'User to connect to the database', { required => 1 } ],
    [ 'pass|p=s', 'Pass to connect to the database', { required => 1 } ],

    [],
    [ 'help|h', 'print usage message and exit' ],
  );

  print $usage->text() && exit if $options->help();

  my $spreadsheet = ReadData( $options->spreadsheet() );
  my $sheet = $spreadsheet->[1];

  my $first_row = 2;
  my $progress = Progress::Any->get_indicator( task => 'upload', target => $sheet->{maxcol} - $first_row );

  chdir( $options->directory() );

  my $dbh = DBI->connect( $options->dsn(), $options->user(), $options->pass() );

  foreach my $line ( $first_row .. $sheet->{maxrow}) {
    my @row = Spreadsheet::Read::row( $sheet, $line );
    $progress->update( message => $row[0] );
    my $id = get_expl_id_by_expl_cb( $dbh, $row[0] );
    # On ajoute un enregistrement dans explnum
    insert_explnum( $dbh, $row[0], $id );
    # On 
    print $row[0], ' --> ', $id, $/;
  }

  $progress->finish();
}

sub get_expl_id_by_expl_cb {
  my ( $dbh, $expl_cb ) = @_;
  my $sql = 'SELECT expl_id FROM exemplaires WHERE expl_cb = ?';
  my $sth = $dbh->prepare( $sql );
  $sth->execute( $expl_cb );
  my $result = $sth->fetchall_arrayref();
  return $result->[0]->[0];
}

sub insert_explnum {
  my ( $dbh, $expl_cb, $expl_id ) = @_;
  my $sql = <<SQL;
  INSERT INTO explnum (explnum_notice, explnum_bulletin, explnum_nom, explnum_mimetype, explnum_url, explnum_data, explnum_vignette, explnum_extfichier, explnum_nomfichier, explnum_statut, explnum_index_sew, explnum_index_wew, explnum_repertoire, explnum_path) VALUES ( ?, 0, ?, 'application/pdf', '', NULL, '', 'pdf', ?, 0, '', '', ?, '/')
SQL
  my $sth = $dbh->prepare( $sql );
  $sth->execute(
    $expl_id,
    $expl_cb . '.pdf',
    $expl_cb . '.pdf',
    1,
  );
}

1;
