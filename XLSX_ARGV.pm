# -*- coding: utf-8 -*-
package XLSX_ARGV;
use strict;
use warnings FATAL => qw/all/;
use sigtrap die => qw(normal-signals);
use Carp;

use File::Spec;
use Archive::Zip qw/:ERROR_CODES/;
use List::Util qw/sum/;

#========================================

sub import {
  my $class = shift;
  if (@_) {
    croak "Unknown arguments: @_";
  }

  $/ = "><";

  tie @main::ARGV, $class, @main::ARGV;
}

#========================================
{
  sub Item () {'XLSX_ARGV::Item'}
  package XLSX_ARGV::Item;
  use fields qw/zip sheets tmpobj/;
  sub new {fields::new(shift)}
}
#========================================

sub new {
  my $self = bless [], shift;
  $self->add_xlsx($_) for @_;
  $self;
}

sub add_xlsx {
  my ($self, $fn) = @_;
  unless (-r $fn) {
    croak "Can't read xlsx file: $fn";
  }
  push @$self, my Item $item = $self->Item->new;
  my $zip = $item->{zip} = Archive::Zip->new;
  unless ((my $rc = $zip->read($fn)) == AZ_OK) {
    croak "Can't open xlsx $fn: return code=$rc\n";
  }
  $item->{sheets} = [$self->list_sheets_from_zip($item->{zip})];
  $self;
}

sub next_file {
  my ($self) = @_;
  my Item $head;
  while (@$self and do {$head = $self->[0]; not @{$head->{sheets}}}) {
    shift @$self;
  }
  return unless $head;
  my $fn = shift @{$head->{sheets}};
  my $tmpobj = $head->{tmpobj} //= File::Temp->newdir;
  my $destfn = File::Spec->catfile($tmpobj->dirname, $fn);
  $head->{zip}->extractMember($fn, $destfn);
  $destfn;
}

sub list_sheets_from_zip {
  my ($self, $zip) = @_;
  map {
    $$_[-1]
  } sort {
    $a->[0] <=> $b->[0]
  } map {
    if ($_->fileName =~ m{^xl/worksheets/sheet(\d+).xml$}) {
      [$1, $&]
    } else {
      ();
    }
  } $zip->members
}

#========================================

sub TIEARRAY {
  shift->new(@_)
}

sub SHIFT {
  shift->next_file
}

sub FETCHSIZE {
  my ($self) = @_;
  sum map {
    my Item $item = $_;
    scalar @{$item->{sheets}}
  } @$self;
}

1;

