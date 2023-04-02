#include <stdio.h>
#include <stdlib.h>

int alternating_length(char *s, char x, char y) {
    int len = 0, which = 0;
    for (int i = 0; s[i]; i++) {
        if (s[i] == x) {
            if (which == 0) {
                len++;
                which = 1;
            } else {
                return 0;
            }
        } else if (s[i] == y) {
            if (which == 1) {
                len++;
                which = 0;
            } else {
                return 0;
            }
        }
    }
    return len;
}

int main(void) {
    int n;
    scanf("%d", &n);
    char s[1002];
    scanf("%s", s);
    int maxLen = 0;
    for (char x = 'a'; x <= 'z'; x++) {
        for (char y = 'a'; y <= 'z'; y++) {
            if (x == y) continue;
            int len = alternating_length(s, x, y);
            if (len > maxLen) maxLen = len;
        }
    }
    if (maxLen == 1) maxLen = 0;
    printf("%d\n", maxLen);
    return 0;
}
