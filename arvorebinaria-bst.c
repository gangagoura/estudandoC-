#include<stdlib.h>
#include<stdio.h>

struct ElementoDaArvoreBinaria {
	int dado;
	struct ElementoDaArvoreBinaria * direita, *esquerda;
};

int menu();
void Inserir(ElementoDaArvoreBinaria **ElementoVarredura, int num);
ElementoDaArvoreBinaria* Buscar(ElementoDaArvoreBinaria **ElementoVarredura, int num);
void Consultar_EmOrdem(ElementoDaArvoreBinaria *);

int main() {
	int op, num, pos, c, res;
	ElementoDaArvoreBinaria *root;
	root = (ElementoDaArvoreBinaria *)malloc(sizeof(ElementoDaArvoreBinaria));
	root = NULL;

	ElementoDaArvoreBinaria *ElementoBusca;
	ElementoBusca = (ElementoDaArvoreBinaria *)malloc(sizeof(ElementoDaArvoreBinaria));

	while (1) {
		op = menu();
		switch (op) {
		case 1:
			printf("Digite o numero desejado: ");
			scanf_s("%d", &num);
			while ((c = getchar()) != '\n' && c != EOF) {} // sempre limpe o buffer do teclado.
			Inserir(&root, num);
			break;
		case 2:
			printf("Digite o numero a ser buscado: ");
			scanf_s("%d", &num);
			while ((c = getchar()) != '\n' && c != EOF) {} // sempre limpe o buffer do teclado.
			ElementoBusca = Buscar(&root, num);
			if (ElementoBusca != 0)
				printf("Valor localizado.\n");
			else
				printf("Valor nao encontrado na arvore.\n");
			system("pause");
			break;
		case 3:
			printf("\n\n");
			Consultar_EmOrdem(root);
			printf("\n\n");
			system("pause");
			break;
		case 4:
			return 0;
		default:
			printf("Invalido\n");
		}
	}
	return 0;
}

int menu() {
	int op, c;
	system("Cls");

	printf("1.Inserir na Arvore Binaria do tipo BST\n");
	printf("2.Buscar na Arvore Binaria do tipo BST\n");
	printf("3.Consultar a Arvore Binaria do tipo BST\n");
	printf("4.Sair\n");
	printf("Digite sua escolha: ");

	scanf_s("%d", &op);
	while ((c = getchar()) != '\n' && c != EOF) {} // sempre limpe o buffer do teclado.

	system("Cls");
	return op;
}

void Inserir(ElementoDaArvoreBinaria ** ElementoVarredura, int num) {
	
	if (*ElementoVarredura == NULL)
	{ //SE ESTÁ VAZIO, COLOCA NA ÁRVORE
		ElementoDaArvoreBinaria *NovoElemento = NULL;
		NovoElemento = (ElementoDaArvoreBinaria *)malloc(sizeof(ElementoDaArvoreBinaria));
		NovoElemento->esquerda = NULL;
		NovoElemento->direita = NULL;

		NovoElemento->dado = num;
		*ElementoVarredura = NovoElemento;
		return;
	}

	if (num < (*ElementoVarredura)->dado) 
	{
		Inserir(&(*ElementoVarredura)->esquerda, num);
	}
	else 
	{
		if (num > (*ElementoVarredura)->dado)
		{
			Inserir(&(*ElementoVarredura)->direita, num);
		}
	}
}

ElementoDaArvoreBinaria* Buscar(ElementoDaArvoreBinaria ** ElementoVarredura, int num)
{
	if (*ElementoVarredura == NULL)
		return NULL;

	if (num < (*ElementoVarredura)->dado)
	{
		Buscar(&((*ElementoVarredura)->esquerda), num);
	}
	else
	{
		if (num > (*ElementoVarredura)->dado)
		{
			Buscar(&((*ElementoVarredura)->direita), num);
		}
		else
		{
			if (num == (*ElementoVarredura)->dado)
				return *ElementoVarredura;
		}
	}
}

void Consultar_EmOrdem(ElementoDaArvoreBinaria * ElementoVarredura) 
{
	if (ElementoVarredura) 
	{
		Consultar_EmOrdem(ElementoVarredura->esquerda);
		printf("%d\t", ElementoVarredura->dado);
		Consultar_EmOrdem(ElementoVarredura->direita);
	}
}
