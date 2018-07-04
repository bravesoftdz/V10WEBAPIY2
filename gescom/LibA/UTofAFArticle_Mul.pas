unit UTofAFArticle_Mul;

interface
uses
{$IFDEF EAGLCLIENT}
     eMul,Maineagl,
{$ELSE}
     Mul, FE_Main, db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}DBGrids,
{$ENDIF}
     Ent1,windows,StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
     HCtrls,HEnt1,HMsgBox,UTOF, UtilGc,
     AfUtilArticle,
     HDimension, DicoBTP, EntGC,UTOFAFTRADUCCHAMPLIBRE;
Type
     TOF_AFARTICLE_MUL = Class (TOF_AFTRADUCCHAMPLIBRE)
        procedure OnUpdate ; override ;
        procedure OnArgument (stArgument : String ) ; override ;
     END ;
Procedure AFLanceFiche_Mul_Article(Range,Argument:string);
Function AFLanceFiche_REch_Article(Argument:string):variant;

implementation
uses UFonctionsCBP;

procedure TOF_AFARTICLE_MUL.OnArgument(stArgument : String );
var
sTmp, sWhereType:string;
stArg:TStringList;
TypeArticle:string;
ComboTypeArticle:THValComboBox;
iPos,iPos1,iPos2:integer;
BEGIN
	fMulDeTraitement := true;
Inherited;
	FTableName := 'ARTICLE';
ComboTypeArticle := nil;

if (Ecran.Name='AFARTICLE_RECH') then
   begin

   // Recuperation de la structure de passage d'arguments
   sTmp := StringReplace(stArgument, ';', chr(VK_RETURN), [rfReplaceAll]);
   stArg:=TStringList.Create;
   try
   stArg.Text:=sTmp;

   // Champ Type Article protégé en ecriture
   if (stArg.Values['TypeArticleEnabled']='false') then
      begin
        SetControlEnabled('GA_TYPEARTICLE', false);
      end;

   // Champ Article fermé initialisé à false
   if (stArg.Values['SansArticleFermes'] = 'true') then
      begin
        SetControlProperty('GA_FERME', 'Checked', false);
      end;

   if (stArg.values['XX_WHERE']<>'') then
      begin
      sWhereType:=stArg.values['XX_WHERE'];
      iPos := pos('GA_TYPEARTICLE', sWhereType);
      iPos1 := pos('(', copy(sWhereType,1,iPos));
      iPos2 := Length (sWhereType);
      if iPos1=0 then inc(iPos1)
      else iPos2 := pos(')', sWhereType);
      if iPos2 > iPos1 then
      begin
        sWhereType := copy (sWhereType,iPos1,iPos2-ipos1+1);
        sWhereType:=StringReplace(sWhereType, 'GA_TYPEARTICLE', 'CO_CODE', [rfReplaceAll]);
        ComboTypeArticle:=THValComboBox(GetControl('GA_TYPEARTICLE'));
        ComboTypeArticle.Plus := 'AND '+ sWhereType;
      end;
            {sWhereType:=StringReplace(sWhereType, 'GA_TYPEARTICLE', 'CO_CODE', [rfReplaceAll]);
      // On doit enlever la condition sur le code CO_CODE="NOM" pour ne pas voir le type nomenclature
      iPosNomenc:=pos('OR CO_CODE="NOM"', sWhereType);
      if (iPosNomenc<>0) then
        begin
        sWhereType := copy(sWhereType, 1, iPosNomenc-1) + copy(sWhereType, iPosNomenc+17, length(sWhereType)-(iPosNomenc+17)+1)
        end;

      ComboTypeArticle:=THValComboBox(GetControl('GA_TYPEARTICLE'));
      ComboTypeArticle.Plus := 'AND '+ sWhereType;    }
      end
   else
      begin
      ComboTypeArticle:=THValComboBox(GetControl('GA_TYPEARTICLE'));
      ComboTypeArticle.plus:=PlusTypeArticle;
      end;

      // PL le 07/09/01 : on bloque toujours le filtre par defaut dans le mul de recherche
      // suite au problème de la recherche sur la saisie d'activite : un filtre sur MAR avait été
      // enregistré par défaut on accédait aux fournitures en saisie par prestations...
      // on pourrait tester l'argument NOFILTRE mais dans le cas de ce mul, aucun filtre ne doit prendre
      // le pas sur des critères par défaut !

   // PL le 03/04/03 : suite à la remarque de MCF, on doit pouvoir mettre un filtre dans les éditions ou la saisie de piece
   // on remet le test
   if stArg.IndexOf('NOFILTRE')<>-1 then
        Tfmul(Ecran).FiltreDisabled:=true;
      // fin PL le 07/09/01

   ComboTypeArticle.value := '';

   // Changement du titre suivant le type de l'article
   TypeArticle:=stArg.Values['GA_TYPEARTICLE'];
   If (TypeArticle<>'') Then
      begin
      	ComboTypeArticle.value:=TypeArticle; //GM
      If (TypeArticle ='PRE') Then
   	   Begin
         Ecran.Caption :='Recherche sur les prestations';
   	   End
      else
      If (TypeArticle ='PRI') Then
   	   Begin
         Ecran.Caption :='Recherche sur les primes';
   	   End
      else
      If (TypeArticle ='POU') Then
   	   Begin
         Ecran.Caption :='Recherche sur les pourcentages';
   	   End
      else
      If (TypeArticle ='FRA') Then
   	   Begin
         Ecran.Caption :='Recherche sur les frais';
   	   End
      else
      If (TypeArticle ='MAR') Then
   	   Begin
         Ecran.Caption :='Recherche sur les fournitures';
   	   End;
      end
   else begin
             //mcd 06/03/02
        end;

   finally
   stArg.free;
   end;
   end;

if (Ecran.Name='AFARTICLE_MUL') then
   begin
   If (stArgument ='PRE') Then
   	Begin
      Ecran.Caption :='Prestations';
   	End
   else
   If (stArgument ='FRA') Then
   	Begin           //mcd 14/04/2003 ??? manque les frais ???
      Ecran.Caption :='Frais';
   	End
   else
   If (stArgument ='PRI') Then
   	Begin
      Ecran.Caption :='Primes';
   	End
   else
   If (stArgument ='CTR') Then
   	Begin
      Ecran.Caption :='Articles de contrat';
   	End
   else
   If (stArgument ='POU') Then
   	Begin
      Ecran.Caption :='Articles de pourcentage';
   	End
   else
   If (stArgument ='MAR') Then
   	Begin
      Ecran.Caption :='Fournitures';
   	End;
   end;
if Not ExJaiLeDroitConcept(TConcept(gcArtCreat),False) then
  BEGIN
  SetControlVisible('BINSERT',False) ;
  SetControlVisible('B_DUPLICATION',False) ;
  END ;
Updatecaption(Ecran);
{$ifdef CCS3}
   SetControlVisible('PZONE',False) ;
{$ENDIF}
END;


procedure TOF_AFARTICLE_MUL.OnUpdate;
var  F : TFMul ;
    i:integer;
    CC:THLabel ;
begin
inherited ;
if not (Ecran is TFMul) then exit;
F:=TFMul(Ecran) ;

for i:=1 to 3 do
    begin
    CC:= THLabel(GetControl('TGA_FAMILLENIV'+InttoStr(i)));
    CC.Caption:=RechDom('GCLIBFAMILLE','LF'+InttoStr(i),false);
    end;
{$IFDEF EAGLCLIENT}
for i:=0 to F.FListe.ColCount-1 do
    begin
    if copy(F.FListe.Cells[i,0],1,14)='Famille niveau' then
      begin
      CC:=THLabel(F.FindComponent('TGA_FAMILLENIV'+copy(F.FListe.Cells[i,0],16,1))) ;
      if Not CC.Visible then F.Fliste.colwidths[i]:=0 ;
      F.Fliste.cells[i,0]:=CC.Caption;
      end;
    end;
{$ELSE}
for i:=0 to F.FListe.Columns.Count-1 do
    begin
    if copy(F.FListe.Columns[i].Title.caption,1,14)='Famille niveau' then
        begin
        CC:=THLabel(GetControl('TGA_FAMILLENIV'+copy(F.FListe.Columns[i].Title.caption,16,1))) ;
        F.Fliste.columns[i].visible:=CC.visible ;
        F.Fliste.columns[i].Field.DisplayLabel:=CC.Caption ;
        end;
    end;
{$ENDIF}
end;

Procedure AFLanceFiche_Mul_Article(Range,Argument:string);
begin
AGLLanceFiche('AFF','AFARTICLE_MUL',range,'',argument);
end;

Function AFLanceFiche_REch_Article(Argument:string):variant;
begin
result:=AGLLanceFiche('AFF','AFARTICLE_RECH','','',argument);
end;


Initialization
registerclasses([TOF_AFARTICLE_MUL]);
end.
