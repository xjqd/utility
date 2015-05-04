#!/usr/bin/python
def print_recurse(items=None):
    for each_item in items:
        if isinstance(each_item, list):
            print_recurse(each_item)
        else:
            print (each_item)
