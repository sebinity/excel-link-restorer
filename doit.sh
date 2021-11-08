#!/bin/bash

# doit.sh
# 2021-10-07, R.Weber: Search and replace text within compressed files
# 2021-10-12, S.Weber: Edits to be able to replace xlsx/xlsb files

# be sure to install the package fuse-zip from the repostitory
# beforehand, run
# sudo ln -s /proc/self/mounts /etc/mtab

# create an temporary directory for using as mountpoint
TEMP_PATH=$(mktemp -d -p ~/)



# search for every compressed file (in this case Excel files) in the current directory
for z in *.xl*
do
  # mount the compressed file to the temporary directory
  fuse-zip ${z} ${TEMP_PATH}
  # search and replace text inside files of the temporary directory
  find ${TEMP_PATH} -type f -exec perl -i -p -e 's/(\<Relationship\ Id\=\"rId2\".*?\/\>)(\<Relationship\ Id\=\"rId1\".*?Target=")(.*[\\|\/])?(.*?xls.)(.*?)?(\".*?\/\>)/$2$4$6/' {} +
  # unmount the compressed file
  fusermount -u -z ${TEMP_PATH}
done

rmdir ${TEMP_PATH}
