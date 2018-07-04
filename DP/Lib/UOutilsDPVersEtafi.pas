{***********UNITE*************************************************
Auteur  ...... : LMO
Créé le ...... : 29/01/2008
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
unit UOutilsDPVersEtafi;
////////////////////////////////////////////////////////////////////////////////
interface
////////////////////////////////////////////////////////////////////////////////

uses
   {$IFDEF EAGLCLIENT}
   UTOB,
   {$ELSE}
   {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
   Db,
   {$ENDIF EAGLCLIENT}

   {$IFDEF VER150}
   Variants,
   {$ENDIF}
   Forms,
   HEnt1, hmsgbox, SysUtils, hCtrls, FileCtrl, DPJUROutils, GalPatience;



////////////////////////////////////////////////////////////////////////////////
function ExportXls2059FetG (sGuidDossier_p, sPath_p : string) : boolean ;
function Ecrit2059 (sql, fn:string): boolean ;
////////////////////////////////////////////////////////////////////////////////

implementation

////////////////////////////////////////////////////////////////////////////////


{***********A.G.L.Privé.*****************************************
Auteur  ...... : LMO
Créé le ...... : 29/01/2008
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
function ExportXls2059FetG(sGuidDossier_p, sPath_p : string) : boolean;
var
   s,rs : string;
   i : integer ;     // pPath,
   F : TFPatience;
begin
  result:=false ;

   //Juri-entreprise vers ETAFI Décisiv : export des données au format Excel
   F := FenetrePatience('Juri-entreprise vers ETAFI Décisiv.');

   if sPath_p[length(sPath_p)]<>'\' then sPath_p:=sPath_p+'\' ;
   s := GetNomPer(sGuidDossier_p) ;

   for i:=1 to length(s) do
       if s[i] in AlphaNumeric then rs:=rs+s[i] ;

   F.lCreation.Caption := 'Génération du formulaire 2059F - Personnes morales...';
   F.Invalidate;
   Application.ProcessMessages;

     //2059F - PM
     s:= 'select ANN_FORME, ANN_NOM1, ANN_NOM2, ANN_SIREN, ' +
         'ANL_TTNBTOT, ' +
         'ANN_VOIENO, ANN_VOIENOCOMPL, ANN_VOIENOM, ANN_ALRUE2 , ' +
         '(ANL_TTPCTCAP*100) Detention_PourCent, ' +
         'ANN_ALRUE3, ANN_ALCP, ANN_ALVILLE, PY_CODEISO2 ' +
         'from ANNUAIRE ' +
         'left join ANNULIEN on (ANL_GUIDPER=ANN_GUIDPER  and ANL_GUIDPERDOS="' + sGuidDossier_p + '") ' +
         'left join JUTYPEFONCT on JTF_FONCTION=ANL_FONCTION '+
         'left outer join PAYS on PY_PAYS =ANN_PAYS '+
         'where ANN_PPPM="PM" and ANL_TYPEDOS="STE" and ANL_FONCTION in ("ASS","ACT","ACE","ACA","AGI") and ANL_TTPCTCAP>=0.1' ;

     if not Ecrit2059 (s, sPath_p + rs + ' 2059F - Personnes morales') then
     begin
      if F<>Nil then
      begin
         F.Close;
         F.Free;
      end;

       pgiBox('Une erreur est survenue à la génération du formulaire 2059F - Personnes morales.', 'Juri-entreprise vers ETAFI Décisiv.') ;
       exit ;
     end ;

      F.lCreation.Caption := 'Génération du formulaire 2059F - Personnes physiques...';
      F.Invalidate;
      Application.ProcessMessages;

     //2059F - PP
     s:= 'select ANN_TYPECIV, ANN_NOM1 || || ANN_NOM2 Identite, ANN_NOM3, ' +
         'ANL_TTNBTOT, (ANL_TTPCTCAP*100) Detention_PourCent, ' +
         'ANN_DATENAIS, ANN_NODEPTNAIS, ANN_VILLENAIS, ANN_PAYSNAIS, ' +
         'ANN_VOIENO, ANN_VOIENOCOMPL, ANN_VOIENOM, ANN_ALRUE2 , ' +
         'ANN_ALRUE3, ANN_ALCP, ANN_ALVILLE, PY_CODEISO2 ' +
         'from ANNUAIRE ' +
         'left join ANNULIEN on (ANL_GUIDPER=ANN_GUIDPER  and ANL_GUIDPERDOS="' + sGuidDossier_p + '") ' +
         'left join JUTYPEFONCT on JTF_FONCTION=ANL_FONCTION '+
         'left outer join PAYS on PY_PAYS =ANN_PAYS '+
         'where ANN_PPPM="PP" and ANL_TYPEDOS="STE" and ANL_FONCTION in ("ASS","ACT","ACE","ACA","AGI") and ANL_TTPCTCAP>=0.1' ;
     if not Ecrit2059 (s, sPath_p + rs + ' 2059F - Personnes physiques') then
     begin
      if F<>Nil then
      begin
         F.Close;
         F.Free;
      end;

       pgiBox('Une erreur est survenue à la génération du formulaire 2059F - Personnes physiques.', 'Juri-entreprise vers ETAFI Décisiv.') ;
       exit ;
     end ;

     //2059G
      F.lCreation.Caption := 'Génération du formulaire 2059G - Filiales et participations...';
      F.Invalidate;
      Application.ProcessMessages;

     s:= 'select ANN_FORME, ANN_NOM1, ANN_NOM2, ANN_SIREN, ' +
         '(ANL_TXDETTOTAL) Detention_PourCent, ' +
         'ANN_VOIENO, ANN_VOIENOCOMPL, ANN_VOIENOM, ANN_ALRUE2 , ' +
         'ANN_ALRUE3, ANN_ALCP, ANN_ALVILLE, PY_CODEISO2 ' +
         'from ANNUAIRE ' +
         'left join ANNULIEN on (ANL_GUIDPER=ANN_GUIDPER  and ANL_GUIDPERDOS="' + sGuidDossier_p + '") ' +
         'left outer join PAYS on PY_PAYS =ANN_PAYS '+
         'where ANL_FONCTION = "FIL" and ANL_TXDETTOTAL>=10' ;
     if not Ecrit2059 (s, sPath_p + rs + ' 2059G - Filiales') then
     begin
      if F<>Nil then
      begin
         F.Close;
         F.Free;
      end;

       pgiBox('Une erreur est survenue à la génération du formulaire 2059G - Filiales et participations.', 'Juri-entreprise vers ETAFI Décisiv.') ;
       exit ;
     end ;

      if F<>Nil then
      begin
         F.Close;
         F.Free;
      end;

  pgiBox('Génération terminée.', 'Juri-entreprise vers ETAFI Décisiv.') ;


  result:=true ;
end ;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 30/01/2008
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function Ecrit2059 (sql, fn:string): boolean ;
var i:integer ;
   FF : TextFile ;
   st : string ;
   q:TQuery ;
begin
result:=false ;

q:=opensql(sql, true) ;

AssignFile(FF,fn+'.csv');
{$I-} Rewrite (FF) ; {$I+}
if IoResult<>0 then begin CloseFile(FF) ; exit ; end ;

for i:=0 to q.FieldCount-1 do
st:=st + q.Fields[i].FieldName + ';' ;
Writeln(FF,st) ;

while not q.eof do
begin
  st:='' ;
  for i:=0 to q.FieldCount-1 do
  begin
    if q.Fields[i].DataType in [ftDate, ftDateTime] then
      st:= st + dateToStr(q.Fields[i].AsDateTime) +';'
    else if q.Fields[i].DataType in [ftInteger, ftSmallint, ftLargeint] then
      st:= st + IntToStr(q.Fields[i].AsInteger) +';'
    else if q.Fields[i].DataType in [ftFloat, ftCurrency] then
      st:= st + STRFPOINT(q.Fields[i].AsFloat) +';'
    else
      st:= st + q.Fields[i].AsVariant +';' ;
  end ;
  Writeln(FF,st) ;
  q.Next ;
end ;
CloseFile(FF) ;
ferme(q) ;
result:=true ;
end;

////////////////////////////////////////////////////////////////////////////////
end.
