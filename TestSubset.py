__author__ = 'staaleu'

import xlsx
import re

if __name__ == "__main__":
    exp = re.compile("([a-zA-Z]+)?((\d+)(\+|-(\d+)?))?(:([a-zA-Z]+)?(\d+(\+|-(\d+))?)?)?")
    print(exp.match("A").groups())
    print(exp.match("A1").groups())
    print(exp.match("1").groups())
    print(exp.match("1:5").groups())
    print(exp.match("A1:C5").groups())
    print(exp.match("A1:C1").groups())
    print(exp.match("A1-100:C1+").groups())
