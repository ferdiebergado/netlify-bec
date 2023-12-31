#!/usr/bin/env bash

for ((i = 1; i < 100; i++)); do
  cp ./data/be_test.xlsx ~/Downloads/test/be_test_copy_$i.xlsx
done
