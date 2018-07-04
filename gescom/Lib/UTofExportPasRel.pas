unit UTofExportPasRel;

interface
uses  StdCtrls,Controls,Classes,db,forms,sysutils,dbTables,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOF, mul, DBGrids,UTOM,AGLInit,
      Utob,HDB,Messages,HStatus,Fe_Main,M3FP,Fiche, PasExport;

Type

     TOF_ExportPasRel = Class (TOF)
     private
        PAS :TPASExport ;
        procedure ExporterDocs ;
        function ExporteLaPiece :  boolean ;
     END ;


procedure AGLPasRelExporterDocuments(parms:array of variant; nb: integer ) ;

implementation


function TOF_ExportPasRel.ExporteLaPiece : boolean ;
var LaPiece,NaturepieceG ,Souche : string;
    Numero, IndiceG : Integer ;
    DatePiece:TDateTime ;
Begin
NaturepieceG:=TFmul(Ecran).Q.FindField('GP_NATUREPIECEG').asstring ;
DatePiece:=TFmul(Ecran).Q.FindField('GP_DATEPIECE').asDateTime ;
Souche:=TFmul(Ecran).Q.FindField('GP_SOUCHE').asstring ;
Numero:=TFmul(Ecran).Q.FindField('GP_NUMERO').asInteger ;
IndiceG:=TFmul(Ecran).Q.FindField('GP_INDICEG').asInteger ;
LaPiece:=NaturepieceG+';'+FormatDateTime('dd/mm/yyyy',DatePiece)+';'+Souche+';'+IntToStr(Numero)+';'+IntToStr(IndiceG)+';;' ;
result:=PAS.ExportDoc (LaPiece,True);
end;


//////////////////////// Export des documents///////////////////////////
procedure TOF_ExportPasRel.ExporterDocs ;
var i : integer;
    L : THDBGrid;
    Q : TQuery ;
    F : TFMul ;
    NbDocs:Integer ;
    OkOk : boolean ;
begin
OkOk:=True;
F:=TFMul(Ecran);
if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
   begin
   MessageAlerte('Aucun élément sélectionné');
   exit;
   end;

if HShowMessage('0;Confirmation;Confirmez vous la génération des pièces ?;Q;YN;N;N;','','')<>mrYes then exit;
Q:=TFMul(Ecran).Q ;
L:=TFMul(Ecran).FListe ;

NbDocs:=0 ;

   try
      PAS:=TPASExport.create ;
      Pas.FichierExport:='Pas.Rel\PAS_Export.fac' ;

      if L.AllSelected then
       BEGIN
       InitMove(Q.RecordCount,'');
       Q.First;
       while Not Q.EOF do
         BEGIN
         MoveCur(False);
         if not ExporteLaPiece  then begin OkOk:=false; break;  end;
         inc(NbDocs);
         Q.NEXT;
         END;
         L.AllSelected:=False;
       END ELSE
       BEGIN
       InitMove(L.NbSelected,'');
       for i:=0 to L.NbSelected-1 do
         BEGIN
         L.GotoLeBOOKMARK(i);
         if not ExporteLaPiece  then begin OkOk:=false; break;  end;
         inc(NbDocs);
         MoveCur(False);
         end ;
       L.ClearSelected;
       END;
       if not OkOk then PgiBox('Traitement non terminé - '+inttostr(NbDocs)+' document(s) exporté(s)','Export Pas.Rel')
               else PgiInfo(inttostr(NbDocs)+' document(s) exporté(s)','Export Pas.Rel') ;
   finally
       FiniMove;
       PAS.free;
   end;
End;

/////////////// Procedure appellé depuis la fiche AGL //////////////
procedure AGLPasRelExporterDocuments(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then TOTOF:=TFMul(F).LaTOF else exit;
if (TOTOF is TOF_ExportPasRel) then TOF_ExportPasRel(TOTOF).ExporterDocs else exit;
end;

Initialization
registerclasses([TOF_ExportPasRel]);
RegisterAglProc('PasRelExporterDocuments',True,1,AGLPasRelExporterDocuments);
end.

