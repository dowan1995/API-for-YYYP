#!/bin/bash

set -e

arr=($(awk -F: '{sub(/^[\t ]*/,"");print $1}' goods.py | grep '^[0-9]'))

for i in "${arr[@]}" ; do
    python3 GetTopLeaseOutOrderList.py $i;
done
