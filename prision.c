/* BROUGHT TO YOU BY code-projects.org */
#include<stdio.h>

#include<stdlib.h>

#include<conio.h>

#include<string.h>

void addrecord();

void viewrecord();

void editrecord();

void searchrecord();

void deleterecord();

void login();

struct record

{
	char id[10];
    
    char name[30];
    
    char age[6];
    
    char gender[10];
    
    char weight[20];
    
    char height[20];
    
    char haircolor[20];
    
    char eyecolor[20];
    
    char crime[40];
    
    char convictedf[20];
    
    char court[20];
	
	char lawyer[20]; 
    
    char punishment[50];
    
    char punishments[20];
    
    char punishmente[20];
    
    char emergencyc[20];
    
    char emergencyr[20];
    
    char emergencyn[20];

} a;

int main()


{
	login();

    int ch;
    printf("\n\n\t***********************************\n");
    printf("\t\t*PRISONER RECORD*\n");
    printf("\t***********************************");

    while(1)
    {

        printf("\n\n\t\tMAIN MENU:");
        printf("\n\n\tADD RECORD\t[1]");
        printf("\n\tSEARCH RECORD\t[2]");
        printf("\n\tEDIT RECORD\t[3]");
        printf("\n\tVIEW RECORD\t[4]");
		printf("\n\tDELETE RECORD\t[5]");
        printf("\n\tEXIT\t\t[6]");
        printf("\n\n\tENTER YOUR CHOICE:");
        scanf("%d",&ch);

        switch(ch)
        {

        case 1:
            addrecord();
            break;

        case 2:
            searchrecord();
            break;

        case 3:
            editrecord();
            break;

        case 4:
            viewrecord();
            break;
            
        case 5:
            deleterecord();
            break;

        case 6:
        	system("cls");
            printf("\n\n\t\tTHANK YOU FOR USING THIS SOFTWARE \n\t\t THIS PROJECT IS BROUGHT TO YOU BY code-projects.org ");
            getch();
            exit(0);

        default:
            printf("\nYOU ENTERED WRONG CHOICE.");
            printf("\nPRESS ANY KEY TO TRY AGAIN");
            getch();
            break;

        }

        system("cls");

    }
    return 0;

}


//*********************************************************
/**************** ADDING FUNCTION************************/
//*********************************************************
void addrecord( )

{
    system("cls");
    FILE *fp ;
    char another = 'Y' ,id[10];
    char filename[30];
    int choice;

    printf("\n\n\t\t***************************\n");
    printf("\t\t* WELCOME TO THE ADD MENU *");
    printf("\n\t\t***************************\n\n");
    printf("\n\n\tENTER FIRST NAME OF PRISONER:\t");
    fflush(stdin);
    gets(filename);

    fp = fopen ("filename", "ab+" ) ;

    if ( fp == NULL )
    {

        fp=fopen("filename","wb+");
        if(fp==NULL)
        {

            printf("\nSYSTEM ERROR...");
            printf("\nPRESS ANY KEY TO EXIT");
            getch();
            return ;

        }
    }

    while ( another == 'Y'|| another=='y' )

    {
        choice=0;
        fflush(stdin);
        
		printf ( "\n\tENTER PRISONER ID:\t");
        scanf("%s",id);
        
		rewind(fp);

        while(fread(&a,sizeof(a),1,fp)==1)
        {
            if(strcmp(a.id,id)==0)
            {

                printf("\n\tTHE RECORD ALREADY EXISTS.\n");
                choice=1;

            }

        }

        if(choice==0)
        {

            strcpy(a.id,id);

            printf("\tENTER PRISONER NAME: ");
            fflush(stdin);
            gets(a.name);
            
			printf("\tENTER GENDER: ");
            gets(a.gender);
            fflush(stdin);

            printf("\tENTER AGE: ");
            gets(a.age);
            fflush(stdin);

            printf("\tENTER WEIGHT: ");
            gets(a.weight);
            fflush(stdin);
            
            printf("\tENTER HEIGHT: ");
            gets(a.height);
 			fflush(stdin);
           
            printf("\tENTER HAIRCOLOR: ");
            gets(a.haircolor);
            fflush(stdin);
            
            printf("\tENTER EYECOLOR: ");
            gets(a.eyecolor);
            fflush(stdin);
            
            printf("\tENTER CRIME: ");
            gets(a.crime);
            fflush(stdin);
            
            printf("\tENTER THE PLACE WHERE THE PRISONER WAS CONVICTED: ");
            gets(a.convictedf);
            fflush(stdin);
            
            printf("\tENTER COURT: ");
            gets(a.court);
            fflush(stdin);
            
            printf("\tENTER LAWYER: ");
            gets(a.lawyer);
            fflush(stdin);
            
            printf("\tENTER CONVICTION: ");
            gets(a.punishment);
            fflush(stdin);
            
            printf("\tENTER THE DATE PUNISHMENT STARTED AT: ");
            gets(a.punishments);
            fflush(stdin);
            
            printf("\tENTER THE DATE PUNISHMENT ENDS AT: ");
            gets(a.punishmente);
            fflush(stdin);
            
            printf("\tENTER NAME OF EMERGENCY CONTACT: ");
            gets(a.emergencyc);
            fflush(stdin);
            
            printf("\tENTER RELATION OF EMERGENCY CONTACT WITH PRISONER: ");
            gets(a.emergencyr);
            fflush(stdin);
            
            printf("\tENTER PHONE NUMBER OF EMERGENCY CONTACT: ");
            gets(a.emergencyn);
            
            fwrite ( &a, sizeof ( a ), 1, fp ) ;
            printf("\nYOUR RECORD IS ADDED...\n");

        }

        printf ( "\n\tADD ANOTHER RECORD...(Y/N) \t" ) ;
        fflush ( stdin ) ;
        another = getch( ) ;

    }

    fclose ( fp ) ;
    printf("\n\n\tPRESS ANY KEY TO EXIT...");
    getch();

}


//***********************************************************
/**************** SEARCHING FUNCTION************************/
//***********************************************************
void searchrecord( )

{
	system("cls");
    FILE *fp ;
	char id[16],choice,filename[14];
    int ch;

    printf("\n\n\t\t*******************************\n");
    printf("\t\t* HERE IS THE SEARCHING MENU *");
    printf("\n\t\t*******************************\n\n");

    do
	{
        
		printf("\n\tENTER THE PRISONER NAME TO BE SEARCHED:");
        fflush(stdin);
        gets(filename);

        fp = fopen ( "filename", "rb" ) ;


        //system("cls");
    	
    		printf("\nENTER PRISONER ID:");
            gets(id);
            system("cls");
            printf("\nTHE WHOLE RECORD IS:");
 
            while ( fread ( &a, sizeof ( a ), 1, fp ) == 1 )

          //{
          if(strcmpi(a.id,id)==0)
               { printf("\n");
                printf("\nPRISONER'S NAME IS: %s",a.name);
                printf("\nPRISONER'S GENDER IS: %s",a.gender);
                printf("\nPRISONER'S AGE IS: %s",a.age);
                printf("\nPRISONER'S WEIGHT IS: %s",a.weight);
                printf("\nPRISONER'S HEIGHT IS: %s",a.height);
                printf("\nPRISONER'S HAIRCOLOR IS: %s",a.haircolor);
				printf("\nPRISONER'S EYECOLOR IS: %s",a.eyecolor);
                printf("\nPRISONER'S CRIME IS: %s",a.crime);
                printf("\nTHE PLACE WHERE THE PRISONER WAS CONVICTED IS: %s",a.convictedf);
                printf("\nCOURT IS: %s",a.court);
                printf("\nPRISONER'S LAWYER IS: %s",a.lawyer);
                printf("\nPRISONER'S CONVICTION IS: %s",a.punishment);
                printf("\nTHE DATE PUNISHMENT STARTED AT IS: %s",a.punishments);
                printf("\nTHE DATE PUNISHMENT ENDS AT IS: %s",a.punishmente);
                printf("\nPRISONER'S EMERGENCY CONTACT IS: %s",a.emergencyc);
                printf("\nRELATION OF EMERGENCY CONTACT WITH PRISONER IS: %s",a.emergencyr);
                printf("\nEMERGENCY CONTACT'S PHONE NUMBER IS: %s",a.emergencyn);
                printf("\n");
            }

           // }
            

        printf("\n\nWOULD YOU LIKE TO CONTINUE VIEWING...(Y/N):");
        fflush(stdin);

        scanf("%c",&choice);

    }
    while(choice=='Y'||choice=='y');

    fclose ( fp) ;
	getch();
    return ;
getch();
}


//*********************************************************
/**************** EDITING FUNCTION************************/
//*********************************************************

void editrecord()

{

    system("cls");
    FILE *fp ;
    char id[10],choice,filename[14];
    int num,count=0;

    printf("\n\n\t\t*******************************\n");
    printf("\t\t* WELCOME TO THE EDITING MENU *");
    printf("\n\t\t*******************************\n\n");

    do
    {

        printf("\n\tENTER THE PRISONER NAME TO BE EDITED:");
        fflush(stdin);
        gets(filename);

        printf("\n\tENTER ID:");
    	gets(id);
        fp = fopen ( "filename", "rb+" ) ;

        /*if ( fp == NULL )
        {

            printf( "\nRECORD DOES NOT EXIST:" ) ;
            printf("\nPRESS ANY KEY TO GO BACK");
            getch();
            return;

        }*/

        while ( fread ( &a, sizeof ( a ), 1, fp ) == 1 )
        {

            if(strcmp(a.id,id)==0)
            {

                printf("\nYOUR OLD RECORD WAS AS:");
                printf("\nPRISONER'S NAME: %s",a.name);
                printf("\nPRISONER'S GENDER: %s",a.gender);
                printf("\nPRISONER'S AGE: %s",a.age);
                printf("\nPRISONER'S WEIGHT: %s",a.weight);
                printf("\nPRISONER'S HEIGHT: %s",a.height);
                printf("\nPRISONER'S HAIRCOLOR: %s",a.haircolor);
				printf("\nPRISONER'S EYECOLOR: %s",a.eyecolor);
                printf("\nPRISONER'S CRIME: %s",a.crime);
                printf("\nTHE PLACE WHERE THE PRISONER WAS CONVICTED: %s",a.convictedf);
                printf("\nCOURT: %s",a.court);
                printf("\nPRISONER'S LAWYER: %s",a.lawyer);
                printf("\nPRISONER'S CONVICTION: %s",a.punishment);
                printf("\nTHE DATE PUNISHMENT STARTED AT: %s",a.punishments);
                printf("\nTHE DATE PUNISHMENT ENDS AT: %s",a.punishmente);
                printf("\nPRISONER'S EMERGENCY CONTACT: %s",a.emergencyc);
                printf("\nRELATION OF EMERGENCY CONTACT WITH PRISONER: %s",a.emergencyr);
                printf("\nEMERGENCY CONTACT'S PHONE NUMBER: %s",a.emergencyn);
                
            
                
                printf("\n\n\t\tWHAT WOULD YOU LIKE TO EDIT..");
                
                
                
                printf("\n1.NAME.");
                printf("\n2.GENDER.");
                printf("\n3.AGE.");
                printf("\n4.WEIGHT.");
                printf("\n5.HEIGHT.");
                printf("\n6.HAIRCOLOR.");
                printf("\n7.EYECOLOR.");
                printf("\n8.CRIME.");
                printf("\n9.PLACE WHERE THE PRISONER WAS CONVICTED.");
                printf("\n10.COURT.");
                printf("\n11.LAWYER.");
                printf("\n12.CONVICTION.");
                printf("\n13.STARTING DATE OF PUNISHMENT.");
                printf("\n14.ENDING DATE OF PUNISHMENT.");
                printf("\n15.EMERGENCY CONTACT.");
                printf("\n16.RELATION OF EMERGENCY CONTACT.");
                printf("\n17.EMERGENCY CONTACT'S PHONE NUMBER.");
                printf("\n18.WHOLE RECORD.");
                printf("\n19.GO BACK TO MAIN MENU.");
                
                do
                {

                    printf("\n\tENTER YOUR CHOICE:");
                    fflush(stdin);
                    scanf("%d",&num);
                    fflush(stdin);
                    
                    switch(num)
                    {

                    case 1:
                        printf("\nENTER THE NEW DATA:");
                        printf("\nNAME:");
                        gets(a.name);
                        break;

                    case 2:
                        printf("\nENTER THE NEW DATA:");
                        printf("\nGENDER:");
                        gets(a.gender);
                        break;

                    case 3:
                        printf("\nENTER THE NEW DATA:");
                        printf("\nAGE:");
                        gets(a.age);
                        break;

                    case 4:
                        printf("\nENTER THE NEW DATA:");
                        printf("\nWEIGHT:");
                        gets(a.weight);
                        break;

                    case 5:
                        printf("ENTER THE NEW DATA:");
                        printf("\nHEIGHT:");
                        gets(a.height);
                        break;

                    case 6:
                        printf("ENTER THE NEW DATA:");
                        printf("\nHAIRCOLOR:");
                        gets(a.haircolor);
                        break;

                    case 7:
                        printf("ENTER THE NEW DATA:");
                        printf("\nEYECOLOR:");
                        gets(a.eyecolor);
                        break;

                    case 8:
                        printf("ENTER THE NEW DATA:");
                        printf("\nCRIME:");
                        gets(a.crime);
                        break;

                    case 9:
                        printf("ENTER THE NEW DATA:");
                        printf("\nPLACE:");
                        gets(a.convictedf);
                        break;

                    case 10:
                        printf("ENTER THE NEW DATA:");
                        printf("\nCOURT:");
                        gets(a.court);
                        break;

                    case 11:
                        printf("ENTER THE NEW DATA:");
                        printf("\nLAWYER:");
                        gets(a.lawyer);
                        break;

                    case 12:
                        printf("ENTER THE NEW DATA:");
                        printf("\nCONVICTION:");
                        gets(a.punishment);
                        break;

                    case 13:
                        printf("ENTER THE NEW DATA:");
                        printf("\nSTARTING DATE OF PUNISHMENT:");
                        gets(a.punishments);
                        break;

                    case 14:
                        printf("ENTER THE NEW DATA:");
                        printf("\nENDING DATE OF PUNISHMENT:");
                        gets(a.punishmente);
                        break;

                    case 15:
                        printf("ENTER THE NEW DATA:");
                        printf("\nEMERGENCY CONTACT:");
                        gets(a.emergencyc);
                        break;

                    case 16:
                        printf("ENTER THE NEW DATA:");
                        printf("\nRELATION OF EMERGENCY CONTACT:");
                        gets(a.emergencyr);
                        break;

                    case 17:
                        printf("ENTER THE NEW DATA:");
                        printf("\nPHONE NUMBER OF EMERGENCY CONTACT:");
                        gets(a.emergencyc);
                        break;

                    case 18:
                    	printf("ENTER THE NEW DATA:");
                        printf("\tPRISONER NAME:");
			            gets(a.name);		
						printf("\tGENDER:");
			            gets(a.gender);			
			            printf("\tAGE:");
			            gets(a.age);
			            printf("\tWEIGHT:");
			            gets(a.weight);
			            printf("\tHEIGHT:");
			            gets(a.height);
			            printf("\tHAIRCOLOR:");
			            gets(a.haircolor);
			            printf("\tEYECOLOR:");
			            gets(a.eyecolor);
			            printf("\tCRIME:");
			            gets(a.age);
			            printf("\tTHE PLACE WHERE THE PRISONER WAS CONVICTED:");
			            gets(a.convictedf);
			            printf("\tCOURT:");
			            gets(a.court);
			            printf("\tLAWYER:");
			            gets(a.lawyer);
			            printf("\tCONVICTION:");
			            gets(a.punishment);
			            printf("\tTHE DATE PUNISHMENT STARTED AT:");
			            gets(a.punishments);
			            printf("\tTHE DATE PUNISHMENT ENDS AT:");
			            gets(a.punishmente);
			            printf("\tNAME OF EMERGENCY CONTACT:");
			            gets(a.emergencyc);
			            printf("\tRELATION OF EMERGENCY CONTACT WITH PRISONER:");
			            gets(a.emergencyr);
			            printf("\tPHONE NUMBER OF EMERGENCY CONTACT:");
			            gets(a.emergencyn);
                        break;

                    case 19:
                        printf("\nPRESS ANY KEY TO GO BACK...\n");
                        getch();
                        return ;
                        break;

                    default:
                        printf("\nYOU TYPED SOMETHING ELSE...TRY AGAIN\n");
                        break;

                    }

                }
                while(num<1||num>20);

                fseek(fp,-sizeof(a),SEEK_CUR);

                fwrite(&a,sizeof(a),1,fp);

                fseek(fp,-sizeof(a),SEEK_CUR);

                fread(&a,sizeof(a),1,fp);

                choice=5;

                break;

            }
        }

        if(choice==5)

        {

            system("cls");

            printf("\n\t\tEDITING COMPLETED...\n");
            printf("--------------------\n");
            printf("THE NEW RECORD IS:\n");
            printf("--------------------\n");
            printf("\nPRISONER'S NAME IS: %s",a.name);
            printf("\nPRISONER'S GENDER IS: %s",a.gender);
            printf("\nPRISONER'S AGE IS: %s",a.age);
            printf("\nPRISONER'S WEIGHT IS: %s",a.weight);
            printf("\nPRISONER'S HEIGHT IS: %s",a.height);
            printf("\nPRISONER'S HAIRCOLOR IS: %s",a.haircolor);
			printf("\nPRISONER'S EYECOLOR IS: %s",a.eyecolor);
            printf("\nPRISONER'S CRIME IS: %s",a.crime);
            printf("\nTHE PLACE WHERE THE PRISONER WAS CONVICTED IS: %s",a.convictedf);
            printf("\nCOURT IS: %s",a.court);
            printf("\nPRISONER'S LAWYER IS: %s",a.lawyer);
    		printf("\nPRISONER'S CONVICTION IS: %s",a.punishment);
        	printf("\nTHE DATE PUNISHMENT STARTED AT IS: %s",a.punishments);
            printf("\nTHE DATE PUNISHMENT ENDS AT IS: %s",a.punishmente);
            printf("\nPRISONER'S EMERGENCY CONTACT IS: %s",a.emergencyc);
            printf("\nRELATION OF EMERGENCY CONTACT WITH PRISONER IS: %s",a.emergencyr);
            printf("\nEMERGENCY CONTACT'S PHONE NUMBER IS: %s",a.emergencyn);

            fclose(fp);

            printf("\n\n\tWOULD YOU LIKE TO EDIT ANOTHER RECORD.(Y/N)");
            scanf("%c",&choice);
            count++;

        }
        else
        {

            printf("\nTHE RECORD DOES NOT EXIST::\n");
            printf("\nWOULD YOU LIKE TO TRY AGAIN...(Y/N)");
        	scanf("%c",&choice);

        }
    }
    while(choice=='Y'||choice=='y');

    fclose ( fp ) ;
    printf("\tPRESS ENTER TO EXIT EDITING MENU.");
    getch();

}

//*********************************************************
/**************** VIEWING FUNCTION************************/
//*********************************************************

void viewrecord()

{
    system("cls");
    FILE *fp;
    //struct record a;
    char filename[30];
    
    printf("\n\n\t\t***************************\n");
    printf("\t\t   * LIST OF PRISONERS *");
    printf("\n\t\t***************************\n\n");
    
    fp=fopen("filename","rb");
    rewind(fp);
    while((fread(&a,sizeof(a),1,fp))==1)
    {
    	printf("||NOTE:-PRESS ENTER KEY TO LOAD OTHER RECORDS||\n");
        printf("\nPRISONER'S NAME IS: %s",a.name);
        printf("\nPRISONER'S GENDER IS: %s",a.gender);
        printf("\nPRISONER'S AGE IS: %s",a.age);
        printf("\nPRISONER'S WEIGHT IS: %s",a.weight);
        printf("\nPRISONER'S HEIGHT IS: %s",a.height);
        printf("\nPRISONER'S HAIRCOLOR IS: %s",a.haircolor);
		printf("\nPRISONER'S EYECOLOR IS: %s",a.eyecolor);
        printf("\nPRISONER'S CRIME IS: %s",a.crime);
        printf("\nTHE PLACE WHERE THE PRISONER WAS CONVICTED IS: %s",a.convictedf);
        printf("\nCOURT IS: %s",a.court);
        printf("\nPRISONER'S LAWYER IS: %s",a.lawyer);
        printf("\nPRISONER'S CONVICTION IS: %s",a.punishment);
        printf("\nTHE DATE PUNISHMENT STARTED AT IS: %s",a.punishments);
        printf("\nTHE DATE PUNISHMENT ENDS AT IS: %s",a.punishmente);
        printf("\nPRISONER'S EMERGENCY CONTACT IS: %s",a.emergencyc);
        printf("\nRELATION OF EMERGENCY CONTACT WITH PRISONER IS: %s",a.emergencyr);
        printf("\nEMERGENCY CONTACT'S PHONE NUMBER IS: %s",a.emergencyn);
        getch();
     printf("\n\n");   
        
    }
    fclose(fp);
	getch();
	
}


//*********************************************************
/**************** DELETING FUNCTION************************/
//*********************************************************

void deleterecord( )

{

    system("cls");
    FILE *fp,*ft ;
    struct record file ;
    char filename[15],another = 'Y' ,id[16];;
    int choice,check;
    int j=0;
    char pass[8];

    printf("\n\n\t\t*************************\n");
    printf("\t\t* WELCOME TO DELETE MENU*");
    printf("\n\t\t*************************\n\n");
    
    printf("\nENTER PASSWORD\n");
    int i;
    for(  i=0;i<4;i++)
    {
    	pass[i]=getch();
    	printf("*");
	}
	if (strcmpi(pass,"pass")==0)
	{
	
     printf("\n\t\t*ACCESS GRANTED*\n\n");
   while ( another == 'Y' || another == 'y' )

    {
       
    	printf("\n\tENTER THE NAME OF PRISONER TO BE DELETED:");
		fflush(stdin);
        gets(filename);
        fp = fopen ("filename", "rb" ) ;
        if ( fp == NULL )
                {
                    printf("\nTHE FILE DOES NOT EXIST");
                    printf("\nPRESS ANY KEY TO GO BACK.");
                    getch();
                    return ;
				}
				ft=fopen("temp","wb");
				
				if(ft==NULL)
                {
                	printf("\nSYSTEM ERROR");
                    printf("\nPRESS ANY KEY TO GO BACK");
                    getch();
                    return ;
                }
                printf("\n\tENTER THE ID OF RECORD TO BE DELETED:");
                fflush(stdin);
                gets(id);

        		while(fread(&file,sizeof(file),1,fp)==1)

                {

                    if(strcmp(file.id,id)!=0)

                        fwrite(&file,sizeof(file),1,ft);

                }
        fclose(ft);
        fclose(fp);
        remove("filename");
        rename("temp","filename");
        printf("\nDELETED SUCCESFULLY...");
        getch();

        printf("\n\tDO YOU LIKE TO DELETE ANOTHER RECORD.(Y/N):");
        fflush(stdin);
        scanf("%c",&another);

    }
   

    printf("\n\n\tPRESS ANY KEY TO EXIT...");

    getch();
}
	else
	{
		printf("\nSorry!Invalid password\n");
		exit(0);
	}

}

void login()
{
	int a=0,i=0;
    char uname[10],c=' '; 
    char pword[10],code[10];
    char user[10]="user";
    char pass[10]="pass";
    do
{
	
    printf("\n  **************************  LOGIN FORM  **************************  ");
    printf(" \n                       ENTER USERNAME:-");
	scanf("%s", &uname); 
	printf(" \n                       ENTER PASSWORD:-");
	while(i<10)
	{
	    pword[i]=getch();
	    c=pword[i];
	    if(c==13) break;
	    else printf("*");
	    i++;
	}
	pword[i]='\0';
	//char code=pword;
	i=0;
	//scanf("%s",&pword); 
		if(strcmp(uname,"user")==0 && strcmp(pword,"pass")==0)
	{
	printf("  \n\n\n       WELCOME TO PRISON MANAGEMENT SYSTEM !! YOUR LOGIN IS SUCCESSFUL");
	printf("\n\n\n\t\t\t\tPress any key to continue...");
	getch();//holds the screen
	break;
	}
	else
	{
		printf("\n        SORRY !!!!  LOGIN IS UNSUCESSFUL");
		a++;
		
		getch();//holds the screen
		
	}
}
	while(a<=2);
	if (a>2)
	{
		printf("\nSorry you have entered the wrong username and password for four times!!!");
		
		getch();
		
		}
		system("cls");	
}
