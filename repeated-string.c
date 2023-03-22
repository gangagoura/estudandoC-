#include <stdio.h>
#include <string.h>


int main( void ) {
	char s[1000];
	long long n;
	
	scanf("%s %lld", s, &n);
	
	int m = strlen(s);
	
	long long count = 0;
	for( int i = 0; i < m; i++ ) {
		if( s[i] == 'a' ) count++;
	}
	
	count *= n / m;
	
	for( int i = 0; i < n % m; i++ ) {
		if( s[i] == 'a' ) count++;
	}
	
	printf("%lld\n", count);
	
	return 0;
}
