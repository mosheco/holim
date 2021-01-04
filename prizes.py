# -*- coding: utf-8 -*-

import math
import sys




if len(sys.argv) < 3:
    print >> sys.stderr, "Usage: python ~/Dropbox/prizes.py <num players> <sum> [with bounties]"
    sys.exit()

if not sys.argv[1].isdigit() or not sys.argv[2].isdigit():
    print >> sys.stderr, "Both arguments must be integers"
    sys.exit()

n = int(sys.argv[1])
s = int(sys.argv[2])

with_bounties = len(sys.argv) == 4
with_bounties_string = with_bounties and "+ באונטים" or ""

if n <= 5:
    print ('*פרס אחד:* %d %s' % (s, with_bounties_string))
elif n <= 9:
    print ('*2 פרסים:*  %d ,%d %s' % (s*0.6, s*0.4, with_bounties_string))
elif n <= 14:
    print ('*3 פרסים:*  %d ,%d ,%d %s' % (s*0.5, s*0.3, s*0.2, with_bounties_string))
elif n <= 22:
    print ('*4 פרסים:*  %d ,%d ,%d ,%d %s' % (math.ceil(s*0.42 - 0.1), math.floor(s*0.28 + 0.1), math.ceil(s*0.17 - 0.1), math.floor(s*0.13 + 0.1), with_bounties_string))
elif n <= 50:
    print ('*5 פרסים:*  %d ,%d ,%d ,%d ,%d %s' % (s*0.38, s*0.24, s*0.16, s*0.12, s*0.10, with_bounties_string))

    
