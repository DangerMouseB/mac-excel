// Copyright (c) 2020 David Briant. All rights reserved.


module addin;

import core.stdc.string : strlen, strcpy;
import std.conv : to;
import std.string : toStringz;
import core.stdc.stdlib : malloc, free;


// P - pointer
// C - char*
// D - double
// L - long (64bit) - vba LongLong (signed)
// I - int (32 bit) - vba Long

extern (C) double addDD_D(double a, double b) {return a + b;}


extern (C) long strlenC_L(char* x) {return strlen(x);}


extern (C) char* concatCC_C(char* a, char*b) {
    string _a = to!string(a);
    string _b = to!string(b);
    string t = _a ~ _b;
    char* buf = cast(char*) malloc(t.length + 1);    // need some memory that D GC won't try to collect
    strcpy(buf, toStringz(t));
    return buf;
}


// utilities for marshalling a string owned here into excel
extern (C) void freeP(void* p) {free(p);}
extern (C) void strcpyCC(char* src, char* dest) {strcpy(dest, src);}

