用SetWindowSubclass来子类化，别用SetWindowLong（这个过时了）。
用SetWindowSubclass更安全，而且它正好多出两个自定义参数（你可以自己任意指定它值）。
