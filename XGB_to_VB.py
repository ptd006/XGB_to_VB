# Dump XGBoost tree to VB function
import numpy as np
import xgboost as xgb
import re

def string_parser_VB(s):
    # The basic cases are whether this is a leaf node or internal
    if len(re.findall(r":leaf=", s)) == 0:
        # internal node
        out  = re.findall(r"[\w.-]+", s)
        tabs = re.findall(r"[\t]+", s)

        if len(tabs) > 0:
            tabs1 = re.findall(r"[\t]+", s)[0].replace('\t', '    ')
            return (tabs1 + '        If state = ' + out[0] + ' Then\n' +
                    tabs1 + '            If ' + out[1] +" < " + out[2] + ' Then\n' +
                    tabs1 + '                state = ' + out[4] + '\n' +
                    tabs1 + '            Else\n' +
                    tabs1 + '                state = ' + out[6] + '\n' +
                    tabs1 + '            End If\n' +
                    tabs1 + '        End If\n\n' )
        
        else:
            tabs1 = ''
            return (tabs1 + '        If state = ' + out[0] + ' Then\n' +
                    tabs1 + '            If ' + out[1] +" < " + out[2] + ' Then\n' +
                    tabs1 + '                state = ' + out[4] + '\n' +
                    tabs1 + '            Else\n' +
                    tabs1 + '                state = ' + out[6] + '\n' +
                    tabs1 + '            End If\n' +
                    tabs1 + '        End If\n\n' )    
    else:
        # This is a leaf node and we should return a value
        # out = re.findall(r"[\d.-]+", s) 
        out = re.findall(r"-?[\d.]+(?:e-?\d+)?", s) # note new regexp for e-0# etc
        tabs1 = re.findall(r"[\t]+", s)[0].replace('\t', '    ') + '        '
        return (tabs1 + 'If state = ' + out[0] + ' then\n' +
                tabs1 + '    xgb_tree = ' + out[1] + '\n' + 
                tabs1 + '    Exit Function' + '\n' + 
                tabs1 + 'End If' + '\n\n')


def tree_parser_VB(tree, i):
    if i == 0:
        return ('    If num_booster = 0 Then\n        state = 0\n'
             + "".join([string_parser_VB(tree.split('\n')[i]) for i in range(len(tree.split('\n'))-1)]))
    else:
        return ('    ElseIf num_booster = '+str(i)+' Then\n        state = 0\n'
             + "".join([string_parser_VB(tree.split('\n')[i]) for i in range(len(tree.split('\n'))-1)])) 

def model_to_VB(model, out_file, num_classes=2):
    trees = model.get_dump()
    
    features  = sorted(list( set(re.findall(r"\[(.+)<[0-9\.]+\]+", "".join(trees) ) ) ) )
    features = ",".join(features)
    result = ["Public Function xgb_tree("+features+",num_booster) As Double\n"
                 +"Dim State As Integer\n"]

    for i in range(len(trees)):
        result.append(tree_parser_VB(trees[i], i))
    
    result.append("    End If ' num_booster \n")
    result.append("End Function\n")
    
    class_proto = ('' if num_classes==2 else ',c')
    
    n_trees = len(trees)
    class_tree_i = ('i' if num_classes==2 else str(num_classes) + '*i + c' )
    n_trees_class = int(n_trees if num_classes==2 else n_trees/num_classes )

    with open(out_file, 'w') as the_file:
        the_file.write("".join(result)
                 + "\nPublic Function xgb_predict("+features+class_proto+")  As Double \n"
                 + "  xgb_predict = 0\n"
                 + "  Dim i As Integer\n"
                 + "  For i = 0 to " + str(n_trees_class - 1) + "\n"
                 + "    xgb_predict = xgb_predict + xgb_tree("+features+","+class_tree_i+")\n"
                 + "  Next i"
                 + "\nEnd Function")

import os
os.chdir('/home/peter/ml/xgb_to_VB')

bst = xgb.Booster({'nthread': 4})  # init model
bst.load_model('demo.json')  # load data

model_to_VB(bst, 'demo.txt',3)


# Generate test data (for copy-pasting in to sheet)
from sklearn import datasets
iris = datasets.load_iris()
X = iris.data
import pandas as pd
pd.DataFrame(bst.predict(xgb.DMatrix(X),output_margin=True)).to_clipboard(index=False)
pd.DataFrame(bst.predict(xgb.DMatrix(X))).to_clipboard(index=False)


# Reg exp fix
s = '3:leaf=-9.95066429e-09'
out = re.findall(r"[\d.-]+", s)
out2 = re.findall(r"-?[\d.]+(?:e-?\d+)?", s)
out2,out

# All values are with 0.000001 of each other

