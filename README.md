# XLSX2ERG

This tool allows to create workouts in Excel/[Libre|Open]Office which then get 
converted to `erg` files. `erg` files can be pushed directly to the Wahoo 
Element and probably also to other bike computers.

## Usage

Create workouts to your heart's content. There are already quite a few examples.
Then, adjust your FTP in the `Overview` worksheet.
The following command will convert the XLSX workbooks to `erg` files.

```
cargo run -- <xlsx_file>
```

I copy them to my Wahoo with 
```
aft-mtp-mount ~/mnt
rsync -av -P /path/to/erg/files>/*.erg ~/mnt/USB\ storage/plans/
umount ~/mnt
```

For convenience, I've put this in a function in my ``.zshrc`` so that I just 
call ``workout`` which makes this process effortless.

## Why

For analysis, I use 
[GoldenCheetah](https://github.com/GoldenCheetah/GoldenCheetah/). 
However, it still lacks a planning feature, which will hopefully be added soon. 
I didn't find any open source planning tool which lets me plan and sync 
workouts to my Element Bolt. 
