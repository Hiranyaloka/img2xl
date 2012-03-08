#!/usr/bin/env perl
use strict;
use warnings;
use File::Find;
use File::Spec;
use String::Dirify;
use Image::Magick;
use Spreadsheet::WriteExcel;
use Cwd qw(getcwd chdir);

if (@ARGV == 0)
  {
  print "\nPrepares the graphics files from the folders passed on the command line.\n".
    "Two thumbnails are created from each original and written to current directory.\n".
    "Filenames are dirified and appended with a 3-digit gallery identifier.\n".
    "Only .jpg and .gif files are processed. Multiple folders may be supplied.\n".
    "Creates a galleries.xls spreadsheet in the current working directory, and embeds\n".
    "the thumbnails in column A, with filename linking to the larger image in Column B.\n".
    "Additional columns may be provided.\n\n";
  exit (1);
  }

# Define small thumb and larger thumb dimensions
my $thumb_width = 200;
my $thumb_height = 100;
my $large_width = 896;
my $large_height = 896;
my $header_row_height = 30; #(in excel units)
my $thumb_height_excel = int(0.75 * $thumb_height); # convert thumb height to excel height
# Define column headings (first two are thumbnail and filename with link)
my @column_names = ( 'Thumbnail', 'Filename (click for popup)', 'Description', 'Gallery Order');

my $dest_dir = getcwd();
my $worksheet;
# Create the spreadsheet object
my $workbook   = Spreadsheet::WriteExcel->new($dest_dir . '/galleries.xls');
my $format = $workbook->add_format(align =>'center', valign =>'vcenter');

# Write data to the spreadsheet
my $sheet_count= 1;
my $suffix;
my $row_count;
my $dir_count = @ARGV;
for (@ARGV) {
  $suffix = ($dir_count > 1) ? sprintf("%02d", $sheet_count++) : '';
  $worksheet = $workbook->add_worksheet('sheet#' . $suffix);
  $worksheet->set_column(0, $#column_names, $header_row_height);
  my $col_no;
  for (@column_names) {
    $worksheet->write(0, $col_no++, $_)
  }
  $row_count = 1;
  find (\&wanted, $_);
}

print "Finished\n";

# Search for image files and create thumbnails
sub wanted {
  my $filename = $File::Find::name;
  return if $filename !~ /\.(?:jpg|jpeg|gif)$/i;
  return if -d $filename;
  my $dir = $File::Find::dir;
  my $name = $_;
  my ($short_name, $ext) = $name =~ /^(.*)(\.(?:jpg|jpeg|gif))$/i;
  my $dirified = String::Dirify->dirify($short_name);
  my $src = Image::Magick->new;
  $src->Read($filename);
  my ($height, $width) = $src->Get ('height', 'width');
  # Create larger thumbnail if image is bigger than defined larger thumb size
  my $large_file = $dirified . '_' . $suffix . $ext;
  my $large_scale = scale ($width, $height, $large_width, $large_height);
  $src->Thumbnail (width => $large_scale * $width, height => $large_scale * $height);
  my $large_destination = File::Spec->catfile($dest_dir, $large_file);
  $src->Write ($large_destination);
  undef $src;

  my $thumb_file = $dirified . '_' . $suffix . '_thumb' . $ext;
  my $thumb_scale = scale ($width, $height, $thumb_width, $thumb_height);
  my $thumb = Image::Magick->new;
  $thumb->Read($large_destination);
  $thumb->Thumbnail (width => $thumb_scale * $width, height => $thumb_scale * $height);
  my $thumb_destination = File::Spec->catfile($dest_dir, $thumb_file);
  $thumb->Write ($thumb_destination);
  undef $thumb;

  my $here = getcwd();
  chdir($dest_dir) unless ($here eq $dest_dir);
  $worksheet->set_row($row_count, $thumb_height_excel, $format);
  $worksheet->insert_image($row_count, 0, $thumb_file);
  $worksheet->write($row_count,1, 'external:' . $large_file);
  $row_count++;
}

# Calculate scale factor based on the proportionately longer side (relative to thumb dimensions)
sub scale {
  my ($width, $height, $thumb_width, $thumb_height) = @_;
  my $lengthX;
  my $lengthY;
  if ($height > $thumb_height or $width > $thumb_width) {
    $lengthX = $thumb_width;
    $lengthY = $thumb_height;
  } else {
    $lengthX = $width;
    $lengthY = $height;
  }

  my $scaleX = $lengthX / $width;
  my $scaleY = $lengthY / $height;
  my $scale = $scaleX < $scaleY ? $scaleX : $scaleY;
}
