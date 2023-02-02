#include <math.h>
#include <stdio.h>
#include <string.h>
#include <stdlib.h>
#include <assert.h>
#include <limits.h>
#include <stdbool.h>

int main(void)
{
    int s, t, a, b, m, n, i, c1 = 0, c2 = 0;
    scanf("%d %d\n", &s, &t);
    scanf("%d %d\n", &a, &b); 
    scanf("%d %d\n", &m, &n);

    for(i = 0; i < m; i++)
    {
        int tmp;
        scanf("%d", &tmp);
        tmp = a + tmp;
        
        if(tmp >= s && tmp <= t)
        {
            c1++;
        }
    }

    for(i = 0; i < n; i++)
    {
        int tmp;
        scanf("%d", &tmp);
        tmp = b + tmp;
        
        if(tmp >= s && tmp <= t)
        {
            c2++;
        }
    }
    
    printf("%d\n%d", c1, c2);
}
