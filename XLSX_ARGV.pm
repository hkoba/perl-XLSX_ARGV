#!/usr/bin/env perl
# -*- coding: utf-8 -*-
package XLSX_ARGV; sub MY () {__PACKAGE__}
use strict;
use warnings FATAL => qw/all/;
use sigtrap die => qw(normal-signals);
use Carp;

use File::Spec;
use Archive::Zip qw/:ERROR_CODES/;
use List::Util qw/sum/;

#========================================
use fields qw/filename zip
	      selected_sheets
	      sheet_list
	      sheet_id_dict
	      sheet_name_dict
	      tmpobj/;

{
  sub SheetInfo () {'XLSX_ARGV::SheetInfo'}
  package XLSX_ARGV::SheetInfo;
  use fields qw/name sheetId r:id/;
  sub new {fields::new(shift)}
}

#========================================

sub import {
  my ($class, $fn, @args) = @_;
  tie @main::ARGV, $class, $fn, (@args ? @args : @main::ARGV);
}

#========================================

sub new {
  my ($class, @init) = @_;
  my MY $self = fields::new($class);
  if (@init) {
    $self->open_xlsx(@init);
  }
  $self;
}

sub open_xlsx {
  (my MY $self, my ($fn, @sheetSpecs)) = @_;
  unless (-r $fn) {
    croak "Can't read xlsx file: $fn";
  }
  my $zip = $self->{zip} = Archive::Zip->new;
  unless ((my $rc = $zip->read($fn)) == AZ_OK) {
    croak "Can't open xlsx $fn: return code=$rc\n";
  }
  $self->{filename} = $fn;
  $self->select_sheets(@sheetSpecs);
  $self;
}

sub next_file {
  (my MY $self) = @_;
  return unless @{$self->{selected_sheets}};
  my $fn = shift @{$self->{selected_sheets}};
  my $destfn = $self->tempfile($fn);
  $self->{zip}->extractMember($fn, $destfn);
  $destfn;
}

sub list_worksheets {
  (my MY $self) = @_;
  map {
    my SheetInfo $si = $_;
    $si->{name}
  } @{$self->{sheet_list}};
}

sub select_sheets {
  (my MY $self, my @sheetSpecs) = @_;
  $self->load_workbook;
  my $selected = $self->{selected_sheets} = [];
  if (@sheetSpecs) {
    foreach my $spec (@sheetSpecs) {
      if (my SheetInfo $si = $self->{sheet_name_dict}{$spec}) {
	push @$selected, $self->_sheet_member($si);
      } elsif ($spec =~ /^\d+$/) {
	push @$selected, $self->_sheet_member($spec);
      } else {
	croak "Invalid sheet spec: $spec";
      }
    }
  } else {
    @$selected = map {
      $self->_sheet_member($_)
    } 1 .. @{$self->{sheet_list}};
  }
}

sub _sheet_member {
  (my MY $self, my $key) = @_;
  my $sheetno = do {
    if (ref $key) {
      my SheetInfo $si = $key;
      $si->{sheetId};
    } else {
      $key;
    }
  };
  "xl/worksheets/sheet$sheetno.xml"
}

sub load_workbook {
  (my MY $self) = @_;
  $self->{sheet_list} = [];
  $self->{sheet_name_dict} = {};
  my $fh = $self->member_fh("xl/workbook.xml");
  local $/ = "><";
  local $_;
  my $line;
  while (<$fh>) {
    chomp;
    if ($line = m{^sheets$} .. m{^/sheets$}) {
      next if $line == 1 or $line =~ /E0$/;
      s/^sheet//;
      my SheetInfo $si = SheetInfo->new;
      while (s{^ ([\w\:]+)="([^\"]*)\"/?}{}) {
	$si->{$1} = $2;
      }
      push @{$self->{sheet_list}}, $si;
      $self->{sheet_id_dict}{$si->{sheetId}} = $si;
      $self->{sheet_name_dict}{$si->{name}} = $si;
    } elsif ($line = m{^definedNames$} .. m{^/definedNames$}) {
      next if $line == 1 or $line =~ /E0$/;
      # XXX:
    } else {
      # discarded.
    }
  }
  wantarray ? @{$self->{sheet_list}} : $self->{sheet_list};
}

#========================================

sub member_fh {
  (my MY $self, my $fn) = @_;
  my $contents = $self->{zip}->contents($fn)
    or croak "Can't find $fn in $self->{filename}";
  open my $fh, '<', \$contents
    or croak "Can't open memory file for $fn: $!";
  $fh;
}

sub tempfile {
  (my MY $self, my $fn) = @_;
  my $tmpobj = $self->{tmpobj} //= File::Temp->newdir;
  File::Spec->catfile($tmpobj->dirname, $fn);
}

#========================================

sub TIEARRAY {
  $/ = "><";
  shift->new(@_)
}

sub SHIFT {
  shift->next_file
}

sub FETCHSIZE {
  (my MY $self) = @_;
  scalar @{$self->{selected_sheets}};
}

#========================================

unless (caller) {
  my @init;
  while (@ARGV) {
    my $arg = shift @ARGV;
    last if $arg eq "--";
    push @init, $arg;
  }
  my MY $self = MY->new(@init);
  my ($method, @rest) = @ARGV;
  $method ||= "list_worksheets";
  require Data::Dumper;
  my @res = $self->$method(@rest);
  foreach my $res (@res) {
    if (ref $res) {
      print Data::Dumper->new([$res])->Terse(1)->Indent(0)->Dump, "\n";
    } else {
      print $res // '', "\n";
    }
  }
}

1;

