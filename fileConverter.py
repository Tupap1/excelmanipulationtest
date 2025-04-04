import os

def changeFilename(path):
    for root, sub, files in os.walk(path):
        for file in files:
            originalFn = file
            replaceFn = originalFn.replace(' ','_')
            oldFn = os.path.join(root, originalFn)
            newFn = os.path.join(root, replaceFn)
            os.rename(oldFn, newFn)
            print(newFn)

    return newFn


changeFilename('SPC TUBING TT04')