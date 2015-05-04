#!/usr/bin/python
def print_recurse(items=None, indent=False, level=0):
    for each_item in items:
        if isinstance(each_item, list):
            print_recurse(each_item, level+1)
        else:
            if indent:
                print("\t" * level)
                #for i in range(levle):
                    #print("\t", end="")
            print (each_item)
