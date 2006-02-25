#include <memory.h>
#include <math.h>
#include <string.h>
#include <stdio.h>
#include <stdlib.h>

#include <libxls/ole.h>

#include <netinet/in.h>

#include <libxls/xlstool.h>

static const DWORD DIFSECT = 0xFFFFFFFC;
static const DWORD FATSECT = 0xFFFFFFFD;
static const DWORD ENDOFCHAIN = 0xFFFFFFFE;
static const DWORD FREESECT = 0xFFFFFFFF;

static int sector_pos(OLE2* ole2, int sid);
static int sector_read(OLE2* ole2, BYTE *buffer, int sid);
static int read_FAT(OLE2* ole2, OLE2Header *oleh);

// Read next sector of stream
extern void ole2_bufread(OLE2Stream* olest)
{
    if (olest->fatpos!=ENDOFCHAIN)
    {
        //printf("Fat val: %X[%X]\n",olest->fatpos,olest->ole->FAT[olest->fatpos]);
        sector_read(olest->ole, olest->buf, olest->fatpos);
        olest->fatpos=olest->ole->FAT[olest->fatpos];
        olest->pos=0;
        olest->cfat++;
    }
}

// Read part of stream
extern int ole2_read(void* buf,long size,long count,OLE2Stream* olest)
{
    int rcount=0;
    DWORD bytes;
    int rem=olest->size-(olest->cfat*olest->ole->lsector+olest->pos);

#if 0
    printf("----------------------------------------------\n");
    printf("ole2_read size=%ld, count=%ld\n", size, count);
#endif

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

#if 0
    printf("----------------------------------------------\n");
    printf("ole2_read (end)\n");
    printf("start:		%li \n",olest->start);
    printf("pos:		%li \n",olest->pos);
    printf("cfat:		%d \n",olest->cfat);
    printf("size:		%d \n",olest->size);
    printf("fatpos:		%li \n",olest->fatpos);
    printf("bufsize:		%li \n",olest->bufsize);
    printf("eof:		%d \n",olest->eof);
#endif

    return(rcount);
}

// Open stream in logical ole file
extern OLE2Stream* ole2_sopen(OLE2* ole,DWORD start)
{
    OLE2Stream* olest=NULL;

#if 0
    printf("----------------------------------------------\n");
    printf("ole2_sopen start=%lXh\n", start);
#endif

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

// Move in stream
extern void ole2_seek(OLE2Stream* olest,DWORD ofs)
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

// Open logical file contained in physical OLE file
extern OLE2Stream* ole2_fopen(OLE2* ole,char* file)
{
    OLE2Stream* olest;
    int i;

#if 0
    printf("----------------------------------------------\n");
    printf("ole2_fopen %s\n", file);
#endif

    for (i=0;i<ole->files.count;i++)
        if (strcmp(ole->files.file[i].name,file)==0)
        {
            olest=ole2_sopen(ole,ole->files.file[i].start);
            olest->size=ole->files.file[i].size;
            return(olest);
        }
    return(NULL);
}

// Open physical file
extern OLE2* ole2_open(char *file)
{
    BYTE buf[1024];
    OLE2Header* oleh;
    OLE2* ole;
    OLE2Stream* olest;
    PSS*	pss;
    char* name = NULL;

#if 0
    printf("----------------------------------------------\n");
    printf("ole2_open %s\n", file);
#endif

    pss=(PSS*)buf;
    oleh=(OLE2Header*) buf;
    ole=(OLE2*)malloc(sizeof(OLE2));
    if (!(ole->file=fopen(file,"rb")))
    {
        //printf("File not found\n");
        free(ole);
        return(NULL);
    }

    // read header and check magic numbers
    fread(buf,1,512,ole->file);

    if (  (ntohl(oleh->id[0]) != 0xD0CF11E0)
        ||(ntohl(oleh->id[1]) != 0xA1B11AE1))
    {
        printf("Not an excel file\n");
        free(ole);
        return(NULL);
    }
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

#if 0
    printf("----------------------------------------------\n");
    printf ("Header Size:	%i \n",sizeof(OLE2Header));
    printf ("verminor:	%lXh  %lXh\n",oleh->id[0],oleh->id[1]);
    printf ("verminor:	%Xh \n",oleh->verminor);
    printf ("verdll:		%Xh \n",oleh->verdll);
    printf ("Byte order:	%Xh \n",oleh->byteorder);
    printf ("sect len:	%Xh (%i)\n",oleh->lsector,ole->lsector);
    printf ("mini len:	%Xh (%i)\n",oleh->lmsector,ole->lmsector);
    printf ("Fat sect.:	%li \n",ole->cfat);
    printf ("Dir Start:	%li \n",ole->dirstart);
    
    printf ("Mini Cutoff:	%li \n",ole->msectorcutoff);
    printf ("MiniFat Start:	%lXh \n",ole->mfatstart);
    printf ("Count MFat:	%li \n",ole->cmfat);
    printf ("Dif start:	%lX \n",ole->difstart);
    printf ("Count Dif:	%li \n",ole->cdif);
    printf ("Fat Size:	%li,%lX \n",ole->cfat*ole->lsector,ole->cfat*ole->lsector);
#endif

    // read directory entries
    read_FAT(ole, oleh);
    olest=ole2_sopen(ole,ole->dirstart);
    do
    {
	// read one directory entry
        ole2_read(pss,1,sizeof(PSS),olest);
#if 0
        		printf("----------------------------------------------\n");
			printf("directory entry\n");
        		printf("name %s\n",utf8_decode(pss->name,sizeof(pss->name), NULL, "iso-8859-1"));
        		printf("bsize %i\n",pss->bsize);
        		printf("type %i\n",pss->type);
        		printf("flag %i\n",pss->flag);
        		printf("left %lX\n",pss->left);
        		printf("right %lX\n",pss->right);
        		printf("child %lX\n",pss->child);
        		printf("guid %.4X-%.4X-%.4X-%.4X %.4X-%.4X-%.4X-%.4X\n",pss->guid[0],pss->guid[1],pss->guid[2],pss->guid[3] ,pss->guid[4],pss->guid[5],pss->guid[6],pss->guid[7]);
        		printf("user flag %.4lX\n",pss->userflags);
        		printf("Start %.4lX\n",pss->sstart);
        		printf("size %.4lX\n",pss->size);
#endif

        // add compound file to list if its name isn't empty
        name=utf8_decode(pss->name, sizeof(pss->name), NULL, "iso-8859-1");
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

// Close physical file
extern void ole2_close(OLE2* ole2)
{
	fclose(ole2->file);
	free(ole2);
}

// Close logical file
extern void ole2_fclose(OLE2Stream* ole2st)
{
	free(ole2st);
}

// Return offset in bytes of a sector from its sid
static int sector_pos(OLE2* ole2, int sid)
{
    return 512 + sid * ole2->lsector;
}

// Read one sector from its sid
static int sector_read(OLE2* ole2, BYTE *buffer, int sid)
{
    fseek(ole2->file, sector_pos(ole2, sid), SEEK_SET);
    fread(buffer, ole2->lsector, 1, ole2->file);
    return 0;
}

// Read FAT
static int read_FAT(OLE2* ole2, OLE2Header* oleh)
{
    int sectorNum;

    // reconstitution of the FAT
    ole2->FAT=malloc(ole2->cfat*ole2->lsector);

    // read first 109 sectors of FAT from header
    {
        int count;
        count = (ole2->cfat < 109) ? ole2->cfat : 109;
        for (sectorNum = 0; sectorNum < count; sectorNum++)
        {
            sector_read(ole2, (BYTE*)(ole2->FAT)+sectorNum*ole2->lsector, oleh->FAT[sectorNum]);
        }
    }

    // Add additionnal sectors of the FAT
    {
        int sid = ole2->difstart;
        BYTE *sector = malloc(ole2->lsector);

        while (sid != ENDOFCHAIN)
          {
           int posInSector;

           // read FAT sector
           sector_read(ole2, sector, sid);

           // read content
           for (posInSector = 0; posInSector < (ole2->lsector-4)/4; posInSector++)
             {
              int s = *(int*)(sector + posInSector*4);
 
              if (s != FREESECT)
                {
                 sector_read(ole2, (BYTE*)(ole2->FAT)+sectorNum*ole2->lsector, s);
                 sectorNum++;
                }
             }

           sid = *(int*)(sector + posInSector*4);
     }

     free(sector);
    }

    return 0;
}

