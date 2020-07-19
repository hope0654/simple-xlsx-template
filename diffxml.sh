#!/bin/bash

root=`pwd`
suffix="dir"

if test $# -ne 2
then
    echo Usage: $0 file1 file2
    exit 1
fi

extract() {
    dir=$2
    rm -rf $dir
    mkdir $dir

    unzip -q $1 -d $dir

    cd $dir

    for file in `find . -name "*.xml"`
    do
        html-beautify -f $file -r
    done

    cd $root
}
dir1="$1-$suffix"
dir2="$2-$suffix"

rm -rf $dir1 $dir2

extract $1 $dir1
extract $2 $dir2

diff -r --exclude workbook.xml --exclude styles.xml --exclude core.xml $dir1 $dir2
