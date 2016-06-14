# -*- coding: UTF-8 -*-

import sys, os, re, gzip

def clean(nin, nout):
    if nin == nout:
        return
    encd = re.compile(r'[ \t]*\#[ \t]*-\*-[ \t]+coding:[ \t]*\S+[ \t]+-\*-[ \t]*\n')
    badt = re.compile(r'(?:(?<=\n)|\A)[ \t]*\n|(?:(?<=\n)|\A)\#[ \t\S]*\n|(?:(?<=\n)|\A)[ \t]*\#[ \t\S]*\n')
    with open(nin) as fin:
        text = fin.read()
    enct = ''
    m = encd.search(text)
    if m:
        enct = m.group(0)
    text = enct + badt.sub('', text)
    with gzip.open(nout, 'wb+') as fout:
        #fout = open(nout, 'w+')
        fout.write(text)
    
def templating(nin, nout, vars):
    import Cheetah.Template
    pvars = {'ConfigurationName' : vars[0], 
            'ProjectName': vars[1],
            'TargetId': vars[1][-2:],
            'PlatformName': vars[2],
            'notdemo': True,
            'demo': True
            }
    try:
        dm = (vars[3].upper() == 'DEMO')
        pvars['demo'] = dm
        pvars['notdemo'] = not dm
    except:
        pass 
    settings = '''
#compiler-settings
directiveStartToken = #$
cheetahVarStartToken = $$
#end compiler-settings
#$if $$ConfigurationName == 'Debug'
    #$set $$debug = True
#$else
    #$set $$debug = False
#$end if\n'''
    if nin == nout:
        return
    with open(nin) as fin:
        text = fin.read()
    text = Cheetah.Template.Template(settings + text.decode('utf-8'), searchList=[pvars])
    with open(nout, 'w+') as fout:
        fout.write(str(text))
    
def zip(nin, nout):
    if nin == nout:
        return
    with open(nin) as fin:
        text = fin.read()
    with gzip.open(nout, 'wb+') as fout:
        fout.write(text)

if __name__ == '__main__':
    if len(sys.argv) == 3:
        clean(sys.argv[1], sys.argv[2])
    elif len(sys.argv) > 3:
        templating(sys.argv[1], sys.argv[2], [v.strip().strip('"') for v in sys.argv[3:]])
