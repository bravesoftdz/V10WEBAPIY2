unit UTofGcUtilisationArtDim;

interface
uses  StdCtrls,Controls,Classes,
{$IFDEF EAGLCLIENT}
			Emul,
{$ELSE}
			db,forms,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
			ComCtrls,DBGrids,mul,
{$ENDIF}
      sysutils,HCtrls,HEnt1,HMsgBox,UTOF,HDimension,Entgc;
Type
     TOF_GcUtilisationArtDim = Class (TOF)
        private
               ListeErreur, Article: String ;
        procedure OnLoad ; override ;
        procedure OnArgument(st: String) ; override ;

     END ;

implementation

procedure TOF_GcUtilisationArtDim.OnArgument(st:String) ;
var ChampMul,ValMul : string;
Critere : string ;
x: integer ;
begin
{récup des arguments }
Repeat
    Critere:=Trim(ReadTokenPipe(st,',')) ;
    if Critere<>'' then
        begin
        x:=pos('=',Critere);
        if x<>0 then
           begin
           ChampMul:=copy(Critere,1,x-1) ;
           ValMul:=copy(Critere,x+1,length(Critere)) ;
           if ChampMul='LISTE' then ListeErreur:=ValMul ;
           if ChampMul='ARTICLE' then Article:=ValMul ;
           end;
        end;
until  Critere='';
end;

procedure TOF_GcUtilisationArtDim.OnLoad;
var i: integer ;
    Texte: String ;
    Liste: TListBox ;
begin
inherited ;
Ecran.Caption:=Article ;
Liste:=TListBox(GetControl('LISTE_UTILISATION')) ;
while ListeErreur <>'' do
  Begin
    i:=StrToInt(ReadTokenSt(ListeErreur)) ;
    Texte:=ReadTokenSt(ListeErreur) ;
    Liste.Items.Add(Texte) ;
  End ;
end;

Initialization
registerclasses([TOF_GcUtilisationArtDim]);
end.
