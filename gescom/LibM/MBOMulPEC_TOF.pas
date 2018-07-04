unit MBOMulPEC_TOF;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,UTOB,
      HCtrls,HEnt1,HMsgBox,UTOF,HDimension,
      Dialogs,M3FP,EntGC,
{$IFDEF EAGLCLIENT}
      emul, MaineAGL,
{$ELSE}
      Fiche, mul, dbTables,Fe_Main,
{$ENDIF}
      LookUp,AglInitGC,UtilGC,UtilArticle,utilDimArticle ;

Type
     TOF_MBOMULPEC = Class (TOF)
       Private
        StatutArt,CodeArticle : string ;
       Public
        procedure OnArgument(Arguments : String) ; override ;
        procedure OnLoad ; override ;

     end ;

implementation


{ TOF_GCMULPEC }

procedure TOF_MBOMULPEC.OnArgument(Arguments: String);
var Critere,ChampMul,ValMul : string;
    x : integer ;
begin
Inherited ;

StatutArt:='' ; CodeArticle:='' ;

repeat
    Critere:=uppercase(Trim(ReadTokenSt(Arguments))) ;
    if Critere<>'' then
        BEGIN
        x:=pos('=',Critere) ;
        if x<>0 then
           BEGIN
           ChampMul:=copy(Critere,1,x-1) ;
           ValMul:=copy(Critere,x+1,length(Critere)) ;

           if ChampMul='GA_STATUTART' then StatutArt:=ValMul
           else if ChampMul='GA_CODEARTICLE' then CodeArticle:=ValMul ;
           END ;
        END ;
until Critere='' ;

end;

procedure TOF_MBOMULPEC.OnLoad;
var xx_where,stStatutArt : string ;
begin
if not (Ecran is TFMul) then exit ;

if (ctxMode in V_PGI.PGIContexte) then
    BEGIN
    {if StatutArt='GEN' then stStatutArt:='DIM' else stStatutArt:=StatutArt ;
    xx_where:='GL_ARTICLE in (select ga_article from article where ga_statutart="'+
              stStatutArt+'" and ga_codearticle="' +CodeArticle+'")' ;}
    if StatutArt='GEN' then // CodeArticle = GA_CODEARTICLE
      xx_where:='GL_ARTICLE in (select ga_article from article where ga_statutart="DIM" ' +
      'and ga_codearticle="' +CodeArticle+'")'
    else // CodeArticle = GA_ARTICLE
      xx_where:='GL_ARTICLE = "' +CodeArticle+'"';
    SetControlText('XX_WHERE',xx_where) ;
    END ;

end;


Initialization
RegisterClasses([TOF_MBOMULPEC]) ;
end.
