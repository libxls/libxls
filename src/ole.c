//#include <malloc.h>
#include <memory.h>
#include <math.h>
#include <string.h>
#include <stdio.h>
#include <stdlib.h>

#include <libxls/ole.h>

const DWORD DIFSECT = 0xFFFFFFFC;
const DWORD FATSECT = 0xFFFFFFFD;
const DWORD ENDOFCHAIN = 0xFFFFFFFE;
const DWORD FREESECT = 0xFFFFFFFF;

extern void ole2_bufread(OLE2Stream* olest) //заполнение(чтение) буфера  OLE2 потока
{
    if (olest->fatpos!=ENDOFCHAIN)
    {
        //printf("Fat val: %X[%X]\n",olest->fatpos,olest->ole->FAT[olest->fatpos]);
        fseek(olest->ole->file,olest->fatpos*olest->ole->lsector+512,0);
        fread(olest->buf,1,olest->bufsize,olest->ole->file);
        olest->fatpos=olest->ole->FAT[olest->fatpos];
        olest->pos=0;
        olest->cfat++;
    }
}

extern int ole2_read(void* buf,long size,long count,OLE2Stream* olest) //чтение из OLE2 потока
{
    int rcount=0;
    DWORD bytes;
    int rem=olest->size-(olest->cfat*olest->ole->lsector+olest->pos);

    if (olest->size>=0)
    {
        bytes=rem<size*count?rem:size*count;
        if (rem<=0)
            olest->eof=1;
    }
    else
        bytes=size*count;

    if (!olest->eof)
    {
        while ((rcount!=bytes)&&(!olest->eof))
        {
            if ((bytes-rcount)<(olest->bufsize-olest->pos))
            {
                memcpy((BYTE*)buf+rcount,olest->buf+olest->pos,bytes-rcount);
                olest->pos+=bytes-rcount;
                rcount+=bytes-rcount;
            }
            else
            {
                memcpy((BYTE*)buf+rcount,olest->buf+olest->pos,olest->bufsize-olest->pos);
                rcount+=olest->bufsize-olest->pos;
                olest->pos+=olest->bufsize-olest->pos;
                ole2_bufread(olest);
            }
            if  ((olest->fatpos==ENDOFCHAIN)&&(olest->bufsize<=olest->pos))
                olest->eof=1;
        }
    }
    return(rcount);
}

extern OLE2Stream* ole2_sopen(OLE2* ole,DWORD start)	//открытие OLE2 потока в OLE2 файле
{
    OLE2Stream* olest=NULL;
    olest=(OLE2Stream*)malloc(sizeof(OLE2Stream));
    olest->ole=ole;
    olest->buf=malloc(ole->lsector);
    olest->bufsize=ole->lsector;
    olest->pos=0;
    olest->eof=0;
    olest->cfat=-1;
    olest->size=-1;
    olest->fatpos=start;
    olest->start=start;
    ole2_bufread(olest);
    return olest;
}

extern void ole2_seek(OLE2Stream* olest,DWORD ofs) //изменение смещения в OLE2 потоке
{
    ldiv_t div_rez=ldiv(ofs,olest->ole->lsector);
    int i;
    olest->fatpos=olest->start;

    if (div_rez.quot!=0)
    {
        for (i=0;i<div_rez.quot;i++)
            olest->fatpos=olest->ole->FAT[olest->fatpos];
    }

    ole2_bufread(olest);
    olest->pos=div_rez.rem;
    olest->eof=0;
    olest->cfat=div_rez.quot;
    //printf("%i=%i %i\n",ofs,div_rez.quot,div_rez.rem);
}

extern 	OLE2Stream*  ole2_fopen(OLE2* ole,char* file)	//открытие OLE2 потока в OLE2 файле
{
    OLE2Stream* olest;
    int i;
    for (i=0;i<ole->files.count;i++)
        if (strcmp(ole->files.file[i].name,file)==0)
        {
            olest=ole2_sopen(ole,ole->files.file[i].start/*номер первого блока*/);
            olest->size=ole->files.file[i].size/*размер файла*/;
            return(olest);
        }
    return(NULL);
}

extern OLE2* ole2_open(char *file) //открытие OLE2 файла
{
    BYTE buf[1024];
    OLE2Header* oleh;
    OLE2* ole;
    OLE2Stream* olest;
    PSS*	pss;
    char* name = NULL;
    int count,i;

    pss=(PSS*)buf;
    oleh=(OLE2Header*) buf;
    ole=(OLE2*)malloc(sizeof(OLE2));
    if (!(ole->file=fopen(file,"rb")))
    {
        //printf("File not found\n");
        free(ole);
        return(NULL);
    }

    fread(buf,1,512,ole->file);

//    ole->lsector=(int)pow(2,oleh->lsector);
//    ole->lmsector=(int)pow(2,oleh->lmsector);
    ole->lsector=512;
    ole->lmsector=64;

    ole->cfat=oleh->cfat;
    ole->dirstart=oleh->dirstart;
    ole->msectorcutoff=oleh->msectorcutoff;
    ole->mfatstart=oleh->mfatstart;
    ole->cmfat=oleh->cmfat;
    ole->difstart=oleh->difstart;
    ole->cdif=oleh->cdif;
    ole->files.count=0;

    printf ("Header Size:	%i \n",sizeof(OLE2Header));
    printf ("verminor:	%Xh  %Xh\n",oleh->id[0],oleh->id[1]);
    printf ("verminor:	%Xh \n",oleh->verminor);
    printf ("verdll:		%Xh \n",oleh->verdll);
    printf ("Byte order:	%Xh \n",oleh->byteorder);
    printf ("sect len:	%Xh (%i)\n",oleh->lsector,ole->lsector);
    printf ("mini len:	%Xh (%i)\n",oleh->lmsector,ole->lmsector);
    printf ("Fat sect.:	%i \n",ole->cfat);
    printf ("Dir Start:	%i \n",ole->dirstart);
    
    printf ("Mini Cutoff:	%i \n",ole->msectorcutoff);
    printf ("MiniFat Start:	%Xh \n",ole->mfatstart);
    printf ("Count MFat:	%i \n",ole->cmfat);
    printf ("Dif start:	%X \n",ole->difstart);
    printf ("Count Dif:	%i \n",ole->cdif);
    printf ("Fat Size:	%i,%X \n",ole->cfat*ole->lsector,ole->cfat*ole->lsector);

    ole->FAT=malloc(ole->cfat*ole->lsector);

    count=(ole->cfat<109)?ole->cfat:109;
    for (i=0;i<count;i++)
    {
        fseek(ole->file,oleh->FAT[i]*ole->lsector+512,0);
        fread((BYTE*)(ole->FAT)+i*ole->lsector,1,ole->lsector,ole->file);
    }

    olest=ole2_sopen(ole,ole->dirstart);
    do
    {
        ole2_read(pss,1,sizeof(PSS),olest);

        /*		printf("name: %s\n",utf8_decode(pss,pss->bsize,NULL,"cp866"));
        		//printf("bsize %i\n",pss->bsize);
        		printf("type %i\n",pss->type);
        		//printf("flag %i\n",pss->flag);
        		//printf("left %X\n",pss->left);
        		//printf("right %X\n",pss->right);
        		//printf("child %X\n",pss->child);
        		//printf("guid %.4X-%.4X-%.4X-%.4X %.4X-%.4X-%.4X-%.4X\n",pss->guid[0],pss->guid[1],pss->guid[2],pss->guid[3]
        		//,pss->guid[4],pss->guid[5],pss->guid[6],pss->guid[7]);
        		//printf("user flag %.4X\n",pss->userflags);
        		printf("Start %.4X\n",pss->sstart);
        		printf("size %.4X\n",pss->size);
        		printf("----------------------------------------------\n");*/
        name=utf8_decode(pss->name,pss->bsize,NULL,"KOI8-R");
        if (name!=NULL)
        {
            if (ole->files.count==0)
            {
                ole->files.file=malloc(sizeof(struct st_olefiles_data));
            }
            else
            {
                ole->files.file=realloc(ole->files.file,(ole->files.count+1)*sizeof(struct st_olefiles_data));
            }
            ole->files.file[ole->files.count].name=name;
            ole->files.file[ole->files.count].start=pss->sstart;
            ole->files.file[ole->files.count].size=pss->size;
            ole->files.count++;
        }

    }
    while (!olest->eof);
    free(olest);
    return ole;
}

extern void ole2_close(OLE2* ole2)
{
	fclose(ole2->file);
	free(ole2);
}

extern void ole2_fclose(OLE2Stream* ole2st)
{
	free(ole2st);
}

