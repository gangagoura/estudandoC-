#include <stdio.h>
#include <string.h>
#include <math.h>
#include <stdlib.h>

int main() {
    int n, mid, d, i, j, k, *a, count = 0;
    
    scanf("%d %d", &n, &d);
    mid = n / 2;
    a = (int *) malloc(n * sizeof(int));
    
    //scanf("%d %d", &a[0], &a[1]);
    for (i = 0; i < n; i++) {
        scanf("%d", &a[i]);
    }
    
    for (i = 0; i < n - 2; i++) {
        j = 1;
        while (a[j] - a[i] <= d) {
            if (a[j] - a[i] == d) {
                k = j + 1;
                while (a[k] - a[j] <= d) {
                    if (a[k] - a[j] == d) {
                        count++;
                        break;
                    }
                    k++;
                }
            }
            j++;
        }
    }
    
    printf("%d", count);
    
    return 0;
}
