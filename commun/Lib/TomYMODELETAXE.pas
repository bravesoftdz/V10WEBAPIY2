{***********UNITE*************************************************
Auteur  ...... : Sylvie DE ALMEIDA
Créé le ...... : 12/02/2007
Modifié le ... : 12/02/2007
Description .. : Source TOM de la TABLE : YMODELETAXE (YMODELETAXE)
Mots clefs ... : TOM;YMODELETAXE
*****************************************************************}
Unit TomYMODELETAXE ;

Interface

Uses
     ed_formu,
     StdCtrls,
     Controls,
     Classes,
     HRichEdt,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fiche,
     FichList,
     FE_Main,       // AGLLanceFiche
     DBCTRLS,       // TDBCheckbox
     HDB,           //THDBMultiValComboBox
{$else}
     eFiche,
     eFichGrid,
     eFichList,
     MaineAGL,      // AGLLanceFiche
     DB,            //THMultiValComboBox
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOM,
     UTob ,
     uiutil,
     HTB97,
     Utobdebug ;

{$IFDEF TAXES}

Procedure YYLanceFiche_ModeleTaxe ;

Type
  TOM_YMODELETAXE = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnAfterDeleteRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    private
      CodeModele : string;
      FromSaisie: boolean;
      NombreTaxesModele  : integer;
      //TOBLignes : TOB;
      strCritereMod : string;
      {$IFNDEF EAGLCLIENT}
      YMT_CODE : THDBEdit;
      YMT_LIBELLE : THDBEdit;
      Y_NOMBRE : THDBSpinEdit;
      YCB_FERME : TDBCheckBox;
      YCB_PRINCIPAL : TDBCheckBox;
      YMT_CRI1LOC : THDBValComboBox;
      YMT_CRI2LOC : THDBValComboBox;
      YMT_CRI3LOC : THDBValComboBox;
      YMT_CRI4LOC : THDBValComboBox;
      YMT_CRI5LOC : THDBValComboBox;
      YMT_CAT1DEP : THDBMultiValComboBox;
      YMT_CAT2DEP : THDBMultiValComboBox;
      YMT_CAT3DEP : THDBMultiValComboBox;
      YMT_CAT4DEP : THDBMultiValComboBox;
      YMT_CAT5DEP : THDBMultiValComboBox;
      {$ELSE}
      ZYMT_CODE : THEdit;
      ZYMT_LIBELLE : THEdit;
      Z_NOMBRE : THSpinEdit;
      ZCB_FERME : TCheckBox;
      ZCB_PRINCIPAL : TCheckBox;
      ZYMT_CRI1LOC : THValComboBox;
      ZYMT_CRI2LOC : THValComboBox;
      ZYMT_CRI3LOC : THValComboBox;
      ZYMT_CRI4LOC : THValComboBox;
      ZYMT_CRI5LOC : THValComboBox;
      ZYMT_CAT1DEP : THMultiValComboBox;
      ZYMT_CAT2DEP : THMultiValComboBox;
      ZYMT_CAT3DEP : THMultiValComboBox;
      ZYMT_CAT4DEP : THMultiValComboBox;
      ZYMT_CAT5DEP : THMultiValComboBox;
      {$ENDIF}

      procedure RechercheEnrPourGriserSuppression;
      procedure NombreOnChange(Sender: TObject);
      function  bVerifFerme(strCodeModele: string) : boolean;
      function  strVerifPrincipal(strCodeModele: string) : string;
      function  bVerifPresencePrincipal : boolean;
      function  bPremierModele : boolean;
      function  bVide : boolean;
      procedure strEnabled (strModele : string);
      procedure strGrise (bGrise : boolean);
    end;

{$ENDIF}

Implementation

{$IFDEF TAXES}
Const MessageListe : Array[0..8] of String =// Message d'erreurs
					({0}'Voulez-vous enregistrer les modifications ?',
           {1}'Merci de renseigner un libellé.',
					 {2}'L''enregistrement est inaccessible.',
           {3}'Veuillez définir un modèle principal.',
           {4}'Confirmez-vous le nouveau modèle principal ?.',
           {5}'Un modèle principal ne peut pas être fermé.',
           {6}'Un modèle principal est obligatoire et vous n''en avez pas défini. Merci de définir un modèle principal.',
           {7}'Merci de renseigner un code modèle de taxe.',
           {8}'Le libellé de la taxe doit être unique. Merci de modifier votre saisie.' );

Procedure YYLanceFiche_ModeleTaxe ;
begin
AGLLanceFiche('YY','YYMODELETAXE','','','') ;
end;


procedure TOM_YMODELETAXE.OnNewRecord ;
begin
  Inherited ;
  {$IFNDEF EAGLCLIENT}
  YMT_CRI1LOC.ItemIndex := 1;
  setfield ('YMT_CRI1LOC', 'PAS');
  if bPremierModele then
  begin
    //YCB_PRINCIPAL.Checked := true;
    setfield ('YMT_PRINCIPAL', 'X');
  end;
  {$ELSE}
  ZYMT_CRI1LOC.ItemIndex := 1;
  setfield ('YMT_CRI1LOC',  'PAS');
  if bPremierModele then
  begin
    ZCB_PRINCIPAL.Checked := true;
    setfield ('YMT_PRINCIPAL', 'X');
  end;
  {$ENDIF}
end ;

procedure TOM_YMODELETAXE.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_YMODELETAXE.OnUpdateRecord ;
var
strLib : string;
i : integer;
OQuery : TQuery;
strVal : string;
bVal : Boolean;


  function IsPrincipal: boolean;
  begin
    if FromSaisie then
      Result := (GetControlText('YMT_PRINCIPAL') = 'X')
    else
      Result := (GetField('YMT_PRINCIPAL') = 'X');
  end;

begin
  Inherited ;
  //Vérification des zones
  //Ecran.Caption := 'Modèles de taxes';
  LastError := 0;
  {$IFNDEF EAGLCLIENT}
   if YMT_CODE.text = '' then
   begin
     PGIBox(MessageListe[7],Ecran.Caption);  //message d'erreur
     SetFocusControl('YMT_CODE');
     LastError := 1;
     exit;
   end;

   if YMT_LIBELLE.text = '' then
   begin
     PGIBox(MessageListe[1],Ecran.Caption);  //message d'erreur
     SetFocusControl('YMT_LIBELLE');
     LastError := 1;
     exit;
   end;
  
    if YCB_PRINCIPAL.Checked then
    begin
        strVal := strVerifPrincipal(YMT_CODE.text);
        if ((strVal <> '') and isPrincipal) then
        begin
        if FromSaisie then
        begin
          if PGIAsk(MessageListe[4]) = mrYes then
          begin
            ExecuteSQL('UPDATE YMODELETAXE SET YMT_PRINCIPAL="-" WHERE YMT_CODE="' + strVal + '"');
          end
          else
          begin
          // Décoche PRINCIPAL
          SetControlText('YMT_PRINCIPAL', '-');
          SetField('YMT_PRINCIPAL', '-');
          end
        end
        else
        begin
          if IsPrincipal then
          ExecuteSQL('UPDATE YMODELETAXE SET YMT_PRINCIPAL="-" WHERE YMT_CODE="' + strVal + '"');
        end
        end;
    end;

    if YCB_FERME.Checked then
    begin
        bVal := bVerifFerme(YMT_CODE.text);
        try
        if ( bVal and IsPrincipal) then
        begin
           PGIBox(MessageListe[5],Ecran.Caption);  //message d'erreur
           SetFocusControl('YCB_FERME');
           LastError := 1;
           Ferme (OQuery);
           exit;
        end
        finally
        Ferme (OQuery);
        end;
    end;
    CodeModele := THDBEdit (GetControl('YMT_CODE')).text;
    NombreTaxesModele := THDBSpinEdit (GetControl('YMT_NOMBRE')).value;
    for i:= 1 to NombreTaxesModele do
    begin
      strLib := 'YMT_LIBCAT' + inttostr(i);
      if (THLabel(GetControl(strLib)).Caption = '') then
      begin
           PGIBox(MessageListe[1],Ecran.Caption);  //message d'erreur
           SetFocusControl(strLib);
           LastError := 1;
           exit;
      end;
    end;

  {$ELSE}
   if ZYMT_CODE.text = '' then
   begin
     PGIBox(MessageListe[7],Ecran.Caption);  //message d'erreur
     SetFocusControl('YMT_CODE');
     LastError := 1;
     exit;
   end;
   if ZYMT_LIBELLE.text = '' then
   begin
     PGIBox(MessageListe[1],Ecran.Caption);  //message d'erreur
     SetFocusControl('YMT_LIBELLE');
     LastError := 1;
     exit;
   end;
    if ZCB_PRINCIPAL.Checked then
    begin
        strVal := strVerifPrincipal(ZYMT_CODE.text);
        if ((strVal <> '') and isPrincipal) then
        begin
        if FromSaisie then
        begin
          if PGIAsk(MessageListe[4]) = mrYes then
          begin
            ExecuteSQL('UPDATE YMODELETAXE SET YMT_PRINCIPAL="-" WHERE YMT_CODE="' + strVal + '"');
          end
          else
          begin
          // Décoche PRINCIPAL
          SetControlText('YMT_PRINCIPAL', '-');
          SetField('YMT_PRINCIPAL', '-');
          end
        end
        else
        begin
          if IsPrincipal then
          ExecuteSQL('UPDATE YMODELETAXE SET YMT_PRINCIPAL="-" WHERE YMT_CODE="' + strVal + '"');
        end
        end;
    end;
    
    if ZCB_FERME.Checked then
    begin
        bVal := bVerifFerme(ZYMT_CODE.text);
        try
        if bVal then
        begin
           PGIBox(MessageListe[5],Ecran.Caption);  //message d'erreur
           SetFocusControl('YMT_FERME');
           LastError := 1;
           Ferme (OQuery);
           exit;
        end
        finally
        Ferme (OQuery);
        end;
    end;
    CodeModele := THEdit (GetControl('YMT_CODE')).text;
    NombreTaxesModele := THSpinEdit (GetControl('YMT_NOMBRE')).value;
    for i:= 1 to NombreTaxesModele do
    begin
      strLib := 'YMT_LIBCAT' + inttostr(i);
      if (THLabel(GetControl(strLib)).Caption = '') then
      begin
           PGIBox(MessageListe[1],Ecran.Caption);  //message d'erreur
           SetFocusControl(strLib);
           LastError := 1;
           exit;
      end
    end;
  {$ENDIF}
    if (getfield ('YMT_LIBCAT1') <> '') and
       ((getfield ('YMT_LIBCAT1') = getfield ('YMT_LIBCAT2')) or
       (getfield ('YMT_LIBCAT1') = getfield ('YMT_LIBCAT3')) or
       (getfield ('YMT_LIBCAT1') = getfield ('YMT_LIBCAT4')) or
       (getfield ('YMT_LIBCAT1') = getfield ('YMT_LIBCAT5'))) then
      begin
           PGIBox(MessageListe[8],Ecran.Caption);  //message d'erreur
           SetFocusControl('YMT_LIBCAT1');
           LastError := 1;
           exit;
      end;
    if (getfield ('YMT_LIBCAT2') <> '') and
       ((getfield ('YMT_LIBCAT2') = getfield ('YMT_LIBCAT1')) or
       (getfield ('YMT_LIBCAT2') = getfield ('YMT_LIBCAT3')) or
       (getfield ('YMT_LIBCAT2') = getfield ('YMT_LIBCAT4')) or
       (getfield ('YMT_LIBCAT2') = getfield ('YMT_LIBCAT5'))) then
      begin
           PGIBox(MessageListe[8],Ecran.Caption);  //message d'erreur
           SetFocusControl('YMT_LIBCAT2');
           LastError := 1;
           exit;
      end;
    if (getfield ('YMT_LIBCAT3') <> '') and
       ((getfield ('YMT_LIBCAT3') = getfield ('YMT_LIBCAT1')) or
       (getfield ('YMT_LIBCAT3') = getfield ('YMT_LIBCAT2')) or
       (getfield ('YMT_LIBCAT3') = getfield ('YMT_LIBCAT4')) or
       (getfield ('YMT_LIBCAT3') = getfield ('YMT_LIBCAT5'))) then
      begin
           PGIBox(MessageListe[8],Ecran.Caption);  //message d'erreur
           SetFocusControl('YMT_LIBCAT3');
           LastError := 1;
           exit;
      end;
    if (getfield ('YMT_LIBCAT4') <> '') and
       ((getfield ('YMT_LIBCAT4') = getfield ('YMT_LIBCAT1')) or
       (getfield ('YMT_LIBCAT4') = getfield ('YMT_LIBCAT2')) or
       (getfield ('YMT_LIBCAT4') = getfield ('YMT_LIBCAT3')) or
       (getfield ('YMT_LIBCAT4') = getfield ('YMT_LIBCAT5'))) then
      begin
           PGIBox(MessageListe[8],Ecran.Caption);  //message d'erreur
           SetFocusControl('YMT_LIBCAT4');
           LastError := 1;
           exit;
      end;
    if (getfield ('YMT_LIBCAT5') <> '') and
       ((getfield ('YMT_LIBCAT5') = getfield ('YMT_LIBCAT1')) or
       (getfield ('YMT_LIBCAT5') = getfield ('YMT_LIBCAT2')) or
       (getfield ('YMT_LIBCAT5') = getfield ('YMT_LIBCAT3')) or
       (getfield ('YMT_LIBCAT5') = getfield ('YMT_LIBCAT4'))) then
      begin
           PGIBox(MessageListe[8],Ecran.Caption);  //message d'erreur
           SetFocusControl('YMT_LIBCAT5');
           LastError := 1;
           exit;
      end;
  LastError := 0;

end ;

procedure TOM_YMODELETAXE.OnAfterUpdateRecord ;
begin
  Inherited ;
  RefreshDB;
end ;

procedure TOM_YMODELETAXE.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_YMODELETAXE.OnLoadRecord ;
var
   i : integer;
   strLib  : string;
   strSQL : string;
   OQuery : TQuery ;
begin
  Inherited ;

  RechercheEnrPourGriserSuppression;

  {if strCritereMod = '' then
     strEnabled (getfield ('YMT_CODE'))
  else
    strEnabled (strCritereMod); }

  if strCritereMod = '' then
   strSQL := 'SELECT 1 FROM YMODELECATTYP where YCY_CODEMODELE = "' + getfield ('YMT_CODE') + '"'
  else
   strSQL := 'SELECT 1 FROM YMODELECATTYP where YCY_CODEMODELE = "' + strCritereMod + '"';
   strSQL := strSQL + ' UNION ';
  if strCritereMod = '' then
   strSQL :=  strSQL +'SELECT 1 FROM YMODELECATREG where YCR_CODEMODELE = "' + getfield ('YMT_CODE') + '"'
  else
   strSQL :=  strSQL +'SELECT 1 FROM YMODELECATREG where YCR_CODEMODELE = "' + strCritereMod + '"';

  OQuery := OpenSql (strSQL, TRUE);
  try
  if not OQuery.Eof then
     exit;
  finally
  Ferme(OQuery);
  end;
  //on gère les zones dépendances (fonction grisée ou non en fonction de case à cocher gérée)
{$IFNDEF EAGLCLIENT}
    CodeModele := THDBEdit (GetControl('YMT_CODE')).text;
    NombreTaxesModele := THDBSpinEdit (GetControl('YMT_NOMBRE')).value;

    for i:= 1 to 5 do
    begin
      strLib := 'LIB_CAT' + inttostr(i);
      THLabel(GetControl(strLib)).enabled := true ;
      strLib := 'TYMT_CATDEP' + inttostr(i);
      if i < 5 then
        THLabel(GetControl(strLib)).enabled := true ;
      strLib := 'TYMT_CATLOC' + inttostr(i);
      THLabel(GetControl(strLib)).enabled := true ;
      strlib := 'YMT_LIBCAT' + inttostr(i);
      THDBEdit (GetControl(strlib)).enabled := true ;
      strLib := 'YMT_CAT' + inttostr(i) + 'DEP';
      if i <> NombreTaxesModele then
        THDBMultiValComboBox (GetControl(strlib)).enabled := true
      else
        THDBMultiValComboBox (GetControl(strlib)).enabled := false;
      strLib := 'YMT_CRI' + inttostr(i) + 'LOC';
      THDBValComboBox (GetControl(strlib)).enabled := true ;
    end;
    for i:= NombreTaxesModele + 1 to 5 do
    begin
      strLib := 'LIB_CAT' + inttostr(i);
      THLabel(GetControl(strLib)).enabled := false ;
      strLib := 'TYMT_CATDEP' + inttostr(i);
      if i < 5 then
        THLabel(GetControl(strLib)).enabled := false ;
      strLib := 'TYMT_CATLOC' + inttostr(i);
      THLabel(GetControl(strLib)).enabled := false ;
      strlib := 'YMT_LIBCAT' + inttostr(i);
      THDBEdit (GetControl(strlib)).enabled := false ;
      strLib := 'YMT_CAT' + inttostr(i) + 'DEP';
      THDBMultiValComboBox (GetControl(strlib)).enabled := false ;
      strLib := 'YMT_CRI' + inttostr(i) + 'LOC';
      THDBValComboBox (GetControl(strlib)).enabled := false ;
    end;
{$ELSE}
    //SDA le 02/04/2007 CodeModele := THEdit (GetControl('ZYMT_CODE')).Text;
    CodeModele := THEdit (GetControl('YMT_CODE')).Text;
    NombreTaxesModele := THSpinEdit (GetControl('YMT_NOMBRE')).value;
    

    for i:=  1 to 5 do
    begin
      strLib := 'LIB_CAT' + inttostr(i);
      THLabel(GetControl(strLib)).enabled := true ;
      strLib := 'TYMT_CATDEP' + inttostr(i);
      if i < 5 then
        THLabel(GetControl(strLib)).enabled := true ;
      strLib := 'TYMT_CATLOC' + inttostr(i);
      THLabel(GetControl(strLib)).enabled := true ;
      strlib := 'YMT_LIBCAT' + inttostr(i);
      THEdit (GetControl(strlib)).enabled := true ;
      strLib := 'YMT_CAT' + inttostr(i) + 'DEP';
      if i <> NombreTaxesModele then
        THMultiValComboBox (GetControl(strlib)).enabled := true
      else
        THMultiValComboBox (GetControl(strlib)).enabled := false;
      strLib := 'YMT_CRI' + inttostr(i) + 'LOC';
      THValComboBox (GetControl(strlib)).enabled := true ;
    end;
    for i:= NombreTaxesModele + 1 to 5 do
    begin
      strLib := 'LIB_CAT' + inttostr(i);
      THLabel(GetControl(strLib)).enabled := false ;
      strLib := 'TYMT_CATDEP' + inttostr(i);
      if i < 5 then
        THLabel(GetControl(strLib)).enabled := false ;
      strLib := 'TYMT_CATLOC' + inttostr(i);
      THLabel(GetControl(strLib)).enabled := false ;
      strlib := 'YMT_LIBCAT' + inttostr(i);
      THEdit (GetControl(strlib)).enabled := false ;
      strLib := 'YMT_CAT' + inttostr(i) + 'DEP';
      THMultiValComboBox (GetControl(strlib)).enabled := false ;
      strLib := 'YMT_CRI' + inttostr(i) + 'LOC';
      THDBValComboBox  (GetControl(strlib)).enabled := false ;
    end;

{$ENDIF}

end ;

procedure TOM_YMODELETAXE.RechercheEnrPourGriserSuppression () ;
begin
//Le bouton de suppression est grisé si des liens BD existent
if DS.State in [dsInsert] then
begin
  TButton(GetControl('bDelete')).enabled := False;
  exit;
end;
if ExisteSQL ('SELECT YMODELECATREG.* FROM YMODELECATREG WHERE YMODELECATREG.YCR_CODEMODELE = "'+getfield ('YMT_CODE')+'" UNION SELECT YMODELECATTYP.* FROM YMODELECATTYP WHERE YMODELECATTYP.YCY_CODEMODELE = "'+getfield ('YMT_CODE')+'" ') then
begin
{$IFNDEF EAGLCLIENT}
   TButton(GetControl('bDelete')).enabled := False;
   TButton(GetControl('bInsert')).enabled := True;
{$ELSE}
   SetControlVisible('BDELETE', FALSE);
   SetControlVisible('Binsert', TRUE);
{$ENDIF}
end
else
begin
{$IFNDEF EAGLCLIENT}
   TButton(GetControl('bDelete')).enabled := True;
   TButton(GetControl('bInsert')).enabled := True;
{$ELSE}
   SetControlVisible('BDELETE', TRUE);
   SetControlVisible('Binsert', TRUE);
{$ENDIF}
end
end;

procedure TOM_YMODELETAXE.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_YMODELETAXE.OnArgument ( S: String ) ;
var
  Critere : string;
  x : integer;
  Arg, Val : string;

begin
  Inherited ;
  strCritereMod := '';
  // Gestion des arguments
  repeat


    Critere := uppercase(Trim(ReadTokenSt(S)));
    if Critere <> '' then
    begin
      x := pos('=', Critere);
      if x <> 0 then
      begin
        Arg := copy(Critere, 1, x - 1);
        Val := copy(Critere, x + 1, length(Critere));
        if Arg = 'CODEMOD' then strCritereMod := Val;
      end;
    end;
  until Critere = '';

  if strCritereMod <> '' then
  begin

  {$IFDEF EAGLCLIENT}
    //TFFicheGrid(Ecran).FRange := strCritereMod ;
    TFFicheGrid(Ecran).FArgument := 'YMT_CODE="'+strCritereMod+'"';
  {$ELSE}
    DS.Filtered := True;
    DS.Filter :=  'YMT_CODE=''' + strCritereMod + '''';
  {$ENDIF}
  //TFFicheListe(Ecran).TypeAction := taconsult;
  end;

  FromSaisie  := true;

{$IFNDEF EAGLCLIENT}

   Y_NOMBRE :=  THDBSpinEdit(GetControl('YMT_NOMBRE', true));
   Y_NOMBRE.OnChange := NombreOnChange;
   YCB_FERME := TDBCheckBox(GetControl('YMT_FERME'));
   YCB_PRINCIPAL := TDBCheckBox(GetControl('YMT_PRINCIPAL'));
   YMT_CODE := THDBEdit (GetControl('YMT_CODE'));
   YMT_LIBELLE := THDBEdit (GetControl('YMT_LIBELLE'));
   YMT_CRI1LOC := THDBValcombobox (GetControl('YMT_CRI1LOC'));
   YMT_CRI2LOC := THDBValcombobox (GetControl('YMT_CRI2LOC'));
   YMT_CRI3LOC := THDBValcombobox (GetControl('YMT_CRI3LOC'));
   YMT_CRI4LOC := THDBValcombobox (GetControl('YMT_CRI4LOC'));
   YMT_CRI5LOC := THDBValcombobox (GetControl('YMT_CRI5LOC'));
   YMT_CAT1DEP := THDBMultiValcombobox (GetControl('YMT_CAT1DEP'));
   YMT_CAT2DEP := THDBMultiValcombobox (GetControl('YMT_CAT2DEP'));
   YMT_CAT3DEP := THDBMultiValcombobox (GetControl('YMT_CAT3DEP'));
   YMT_CAT4DEP := THDBMultiValcombobox (GetControl('YMT_CAT4DEP'));
   YMT_CAT5DEP := THDBMultiValcombobox (GetControl('YMT_CAT5DEP'));
{$ELSE}
   Z_NOMBRE :=  THSpinEdit(GetControl('YMT_NOMBRE', true));
   Z_NOMBRE.OnChange := NombreOnChange;
   ZCB_FERME := TCheckBox(GetControl('YMT_FERME'));
   ZCB_PRINCIPAL := TCheckBox(GetControl('YMT_PRINCIPAL'));
   ZYMT_CODE := THEdit (GetControl('YMT_CODE'));
   ZYMT_LIBELLE := THEdit (GetControl('YMT_LIBELLE'));
   ZYMT_CRI1LOC := THValcombobox (GetControl('YMT_CRI1LOC'));
   ZYMT_CRI2LOC := THValcombobox (GetControl('YMT_CRI2LOC'));
   ZYMT_CRI3LOC := THValcombobox (GetControl('YMT_CRI3LOC'));
   ZYMT_CRI4LOC := THValcombobox (GetControl('YMT_CRI4LOC'));
   ZYMT_CRI5LOC := THValcombobox (GetControl('YMT_CRI5LOC'));
   ZYMT_CAT1DEP := THMultiValcombobox (GetControl('YMT_CAT1DEP'));
   ZYMT_CAT2DEP := THMultiValcombobox (GetControl('YMT_CAT2DEP'));
   ZYMT_CAT3DEP := THMultiValcombobox (GetControl('YMT_CAT3DEP'));
   ZYMT_CAT4DEP := THMultiValcombobox (GetControl('YMT_CAT4DEP'));
   ZYMT_CAT5DEP := THMultiValcombobox (GetControl('YMT_CAT5DEP'));
{$ENDIF}

end ;

procedure TOM_YMODELETAXE.OnClose ;
begin

  Inherited ;
  if (bVerifPresencePrincipal) and (not bvide) then
  begin
     PGIBox(MessageListe[6],Ecran.Caption);  //message d'erreur
     SetFocusControl('YMT_CODE');
     LastError := 1;
     exit;
  end;
end ;

procedure TOM_YMODELETAXE.OnCancelRecord ;
begin
   Inherited ;
end ;

procedure TOM_YMODELETAXE.NombreOnChange(Sender: TObject);
//Changement du nombre de taxes du modèle
var
   i, j : integer;
   strLib, strReq, strDep   : string;
   strCode : string;

begin
  Inherited ;
  strCode := getfield ('YMT_CODE');


 //on gère les zones dépendances (fonction grisée ou non en fonction de case à cocher gérée)
{$IFNDEF EAGLCLIENT}
    //CodeModele := THDBEdit (GetControl('YMT_CODE')).text;
    CodeModele := strCode;
    NombreTaxesModele := THDBSpinEdit (GetControl('YMT_NOMBRE')).value;
    for i := 1 to NombreTaxesModele do
    begin
      strReq := '';
      strLib := 'YMT_CAT' + inttostr(i) + 'DEP';
      if (NombreTaxesModele > 1) and (i <> NombreTaxesModele) then
        strReq := ' AND CC_CODE IN (';
      for j := i + 1  to NombreTaxesModele do
      begin
        strDep := 'TX' + inttostr(j);
        strReq :=  strReq + '"'+strDep+'"';
        if j <> NombreTaxesModele then
          strReq :=  strReq + ','
        else
          strReq := strReq + ' )';
      end;

      if strReq <> '' then
        THDBMultiValComboBox(GetControl(strLib)).Plus :=  strReq;
    end;


    for i:= 1 to 5 do
    begin
      strLib := 'LIB_CAT' + inttostr(i);
      THLabel(GetControl(strLib)).enabled := true ;
      strLib := 'TYMT_CATDEP' + inttostr(i);
      THLabel(GetControl(strLib)).enabled := true ;
      strLib := 'TYMT_CATLOC' + inttostr(i);
      THLabel(GetControl(strLib)).enabled := true ;
      strlib := 'YMT_LIBCAT' + inttostr(i);
      THDBEdit (GetControl(strlib)).enabled := true ;
      strLib := 'YMT_CAT' + inttostr(i) + 'DEP';
      if i <> NombreTaxesModele then
        THDBMultiValComboBox (GetControl(strlib)).enabled := true
      else
        THDBMultiValComboBox (GetControl(strlib)).enabled := false;
      strLib := 'YMT_CRI' + inttostr(i) + 'LOC';
      THDBValComboBox (GetControl(strlib)).enabled := true ;
      if (getfield(strLib) = '' ) and (i < NombreTaxesModele + 1) then
      begin
        //THDBValComboBox (GetControl(strlib)).Value := 'Pas de loc.';
        setfield ('YMT_CRI' + inttostr(i) +'LOC',  'PAS');
      end
    end;
    for i:= NombreTaxesModele + 1 to 5 do
    begin
      strLib := 'LIB_CAT' + inttostr(i);
      THLabel(GetControl(strLib)).enabled := false ;
      strLib := 'TYMT_CATDEP' + inttostr(i);
      THLabel(GetControl(strLib)).enabled := false ;
      strLib := 'TYMT_CATLOC' + inttostr(i);
      THLabel(GetControl(strLib)).enabled := false ;
      strlib := 'YMT_LIBCAT' + inttostr(i);
      THDBEdit (GetControl(strlib)).enabled := false ;
      THDBEdit (GetControl(strlib)).Text := '';
      strLib := 'YMT_CAT' + inttostr(i) + 'DEP';
      THDBMultiValComboBox (GetControl(strlib)).enabled := false ;
      THDBMultiValComboBox (GetControl(strlib)).Text := '';
      strLib := 'YMT_CRI' + inttostr(i) + 'LOC';
      THDBValComboBox (GetControl(strlib)).enabled := false ;

      SetControlText(strLib, '');
      //THDBValComboBox (GetControl(strlib)).Text := '';
      if (THDBValComboBox (GetControl(strlib)).Text ='' ) and (i < NombreTaxesModele + 1) then
      begin
        //THDBValComboBox (GetControl(strlib)).ItemIndex := 1;
        setfield ('YMT_CRI' + inttostr(i) +'LOC',  'PAS');
      end
    end;
{$ELSE}
    //CodeModele := THEdit (GetControl('YMT_CODE')).Text;
    CodeModele := strCode;
    NombreTaxesModele := THSpinEdit (GetControl('YMT_NOMBRE')).value;
    for i := 1 to NombreTaxesModele do
    begin
      strReq := '';
      strLib := 'YMT_CAT' + inttostr(i) + 'DEP';
      if (NombreTaxesModele > 1) and (i <> NombreTaxesModele) then
        strReq := ' AND CC_CODE IN (';
      for j := i + 1  to NombreTaxesModele do
      begin
        strDep := 'TX' + inttostr(j);
        strReq :=  strReq + '"'+strDep+'"';
        if j <> NombreTaxesModele then
          strReq :=  strReq + ','
        else
          strReq := strReq + ' )';
      end;

      if strReq <> '' then
      THMultiValComboBox(GetControl(strLib)).Plus :=  strReq;
    end;
    for i:=  1 to 5 do
    begin
      strLib := 'LIB_CAT' + inttostr(i);
      THLabel(GetControl(strLib)).enabled := true ;
      strLib := 'TYMT_CATDEP' + inttostr(i);
      THLabel(GetControl(strLib)).enabled := true ;
      strLib := 'TYMT_CATLOC' + inttostr(i);
      THLabel(GetControl(strLib)).enabled := true ;
      strlib := 'YMT_LIBCAT' + inttostr(i);
      THEdit (GetControl(strlib)).enabled := true ;
      strLib := 'YMT_CAT' + inttostr(i) + 'DEP';
      if i <> NombreTaxesModele then
        THMultiValComboBox (GetControl(strlib)).enabled := true
      else
        THMultiValComboBox (GetControl(strlib)).enabled := false;
      strLib := 'YMT_CRI' + inttostr(i) + 'LOC';
      THValcomboBox (GetControl(strlib)).enabled := true ;
      if (getfield(strLib) = '' )  and (i < NombreTaxesModele + 1) then
      begin
        //THValComboBox (GetControl(strlib)).ItemIndex := 1;
        setfield ('YMT_CRI' + inttostr(i) +'LOC',  'PAS');
      end
    end;
    for i:= NombreTaxesModele + 1 to 5 do
    begin
      strLib := 'LIB_CAT' + inttostr(i);
      THLabel(GetControl(strLib)).enabled := false ;
      strLib := 'TYMT_CATDEP' + inttostr(i);
      THLabel(GetControl(strLib)).enabled := false ;
      strLib := 'TYMT_CATLOC' + inttostr(i);
      THLabel(GetControl(strLib)).enabled := false ;
      strlib := 'YMT_LIBCAT' + inttostr(i);
      THEdit (GetControl(strlib)).enabled := false ;
      THEdit (GetControl(strlib)).Text := '';
      strLib := 'YMT_CAT' + inttostr(i) + 'DEP';
      THMultiValComboBox (GetControl(strlib)).enabled := false ;
      THMultiValComboBox (GetControl(strlib)).Text := '';
      strLib := 'YMT_CRI' + inttostr(i) + 'LOC';
      THValComboBox (GetControl(strlib)).enabled := false ;
      THValComboBox (GetControl(strlib)).text := '';
      if (THDBValComboBox (GetControl(strlib)).Text = '' )  and (i < NombreTaxesModele + 1) then
      begin
        //THValComboBox (GetControl(strlib)).ItemIndex := 1;
        setfield ('YMT_CRI' + inttostr(i) +'LOC',  'PAS');
      end
    end;

{$ENDIF}

end;

function TOM_YMODELETAXE.strVerifPrincipal(strCodeModele: string) : string;
//Vérification modèle principal : un seul modèle principal possible
var
    strSQL, St : string;
    OQuery : TQuery ;
begin
  result := '';
  strSQL := 'SELECT YMODELETAXE.YMT_CODE FROM YMODELETAXE WHERE YMODELETAXE.YMT_PRINCIPAL = "X" AND YMODELETAXE.YMT_CODE <> "' + strCodeModele + '" GROUP BY YMODELETAXE.YMT_CODE';
  OQuery := OpenSql (strSQL, TRUE);
  try
  if not OQuery.Eof then
  begin
    St := OQuery.Fields[0].AsString;
    result := St;
  end;
  finally
  Ferme (OQuery);
  end;

end;

function TOM_YMODELETAXE.bVerifFerme(strCodeModele: string) : boolean;
//Vérification modèle fermé
var
    strSQL : string;
    OQuery : TQuery ;
    i : string;
begin
  result := false;
  strSQL := 'SELECT YMODELETAXE.YMT_PRINCIPAL FROM YMODELETAXE WHERE YMODELETAXE.YMT_CODE = "' + strCodeModele + '"';
  OQuery := OpenSql (strSQL, TRUE);
  try
  if not OQuery.Eof then
  begin
    i := OQuery.Fields[0].AsString;
    if i = 'X' then  //le modèle est principal et ne peut donc pas être fermé
    begin
      result := true;
    end;
  end;
  finally
  Ferme (OQuery);
  end;
end;

function TOM_YMODELETAXE.bVerifPresencePrincipal : boolean;
//Vérification modèle fermé
var
    strSQL : string;
    OQuery : TQuery ;
begin
  result := false;
  strSQL := 'SELECT YMODELETAXE.YMT_CODE FROM YMODELETAXE WHERE YMODELETAXE.YMT_PRINCIPAL = "X"';
  OQuery := OpenSql (strSQL, TRUE);
  try
    if OQuery.Eof then
    begin
      result := true;
    end;
  finally
    Ferme (OQuery);
  end;
end;

function TOM_YMODELETAXE.bVide : boolean;
//Vérification rien en BD
var
    strSQL : string;
    OQuery : TQuery ;
begin
  result := false;
  strSQL := 'SELECT 1 FROM YMODELETAXE';
  OQuery := OpenSql (strSQL, TRUE);
  try
    if OQuery.Eof then
       result := true;
  finally
    Ferme (OQuery);
  end;
end;

function TOM_YMODELETAXE.bPremierModele: boolean;
var
    strSQL : string;
    OQuery : TQuery ;
begin
  result := false;
  strSQL := 'SELECT count(*) AS TOTAL FROM YMODELETAXE';
  OQuery := OpenSql (strSQL, TRUE);
  try
  if OQuery.Fields[0].AsInteger = 0 then
  begin
      result := true;
  end;
  finally
  Ferme (OQuery);
  end;
end;


procedure TOM_YMODELETAXE.strEnabled(strModele: string);
var
    strSQL : string;
    OQuery : TQuery ;
begin
  strSQL := 'SELECT 1 FROM YMODELECATTYP where YCY_CODEMODELE = "' + strModele + '"';
  strSQL := strSQL + ' UNION ';
  strSQL :=  strSQL +'SELECT 1 FROM YMODELECATREG where YCR_CODEMODELE = "' + strModele + '"';
  OQuery := OpenSql (strSQL, TRUE);
  try
  if not OQuery.Eof then
  begin
      //modèle utilisé dans param tout est grisé sauf libellé
      strGrise(false);
  end
  else
    strGrise(true);
  finally
  Ferme (OQuery);
  end;
end;

procedure TOM_YMODELETAXE.strGrise(bGrise: boolean);
begin
{$IFNDEF EAGLCLIENT}

   Y_NOMBRE.Enabled := bGrise;
   //YCB_FERME.Enabled := bGrise;
   //YCB_PRINCIPAL.Enabled := bGrise;
   THLabel(GetControl('LIB_CAT1')).enabled := bGrise ;
   THLabel(GetControl('LIB_CAT2')).enabled := bGrise ;
   THLabel(GetControl('LIB_CAT3')).enabled := bGrise ;
   THLabel(GetControl('LIB_CAT4')).enabled := bGrise ;
   THLabel(GetControl('LIB_CAT5')).enabled := bGrise ;
   THLabel(GetControl('TYMT_CATDEP1')).enabled := bGrise;
   THLabel(GetControl('TYMT_CATDEP2')).enabled := bGrise;
   THLabel(GetControl('TYMT_CATDEP3')).enabled := bGrise;
   THLabel(GetControl('TYMT_CATDEP4')).enabled := bGrise;
   THLabel(GetControl('TYMT_CATDEP5')).enabled := bGrise;
   THLabel(GetControl('TYMT_CATLOC1')).enabled := bGrise;
   THLabel(GetControl('TYMT_CATLOC2')).enabled := bGrise;
   THLabel(GetControl('TYMT_CATLOC3')).enabled := bGrise;
   THLabel(GetControl('TYMT_CATLOC4')).enabled := bGrise;
   THLabel(GetControl('TYMT_CATLOC5')).enabled := bGrise;
   THDBEdit (GetControl('YMT_LIBCAT1')).enabled := bGrise ;
   THDBEdit (GetControl('YMT_LIBCAT2')).enabled := bGrise ;
   THDBEdit (GetControl('YMT_LIBCAT3')).enabled := bGrise ;
   THDBEdit (GetControl('YMT_LIBCAT4')).enabled := bGrise ;
   THDBEdit (GetControl('YMT_LIBCAT5')).enabled := bGrise ;
   YMT_CRI1LOC.Enabled := bGrise;
   YMT_CRI2LOC.Enabled := bGrise;
   YMT_CRI3LOC.Enabled := bGrise;
   YMT_CRI4LOC.Enabled := bGrise;
   YMT_CRI5LOC.Enabled := bGrise;
   YMT_CAT1DEP.Enabled := bGrise;
   YMT_CAT2DEP.Enabled := bGrise;
   YMT_CAT3DEP.Enabled := bGrise;
   YMT_CAT4DEP.Enabled := bGrise;
   YMT_CAT5DEP.Enabled := bGrise;
{$ELSE}
   Z_NOMBRE.Enabled := bGrise;
   //ZCB_FERME.Enabled := bGrise;
   //YCB_PRINCIPAL.Enabled := bGrise;
   THLabel(GetControl('LIB_CAT1')).enabled := bGrise ;
   THLabel(GetControl('LIB_CAT2')).enabled := bGrise ;
   THLabel(GetControl('LIB_CAT3')).enabled := bGrise ;
   THLabel(GetControl('LIB_CAT4')).enabled := bGrise ;
   THLabel(GetControl('LIB_CAT5')).enabled := bGrise ;
   THLabel(GetControl('TYMT_CATDEP1')).enabled := bGrise;
   THLabel(GetControl('TYMT_CATDEP2')).enabled := bGrise;
   THLabel(GetControl('TYMT_CATDEP3')).enabled := bGrise;
   THLabel(GetControl('TYMT_CATDEP4')).enabled := bGrise;
   THLabel(GetControl('TYMT_CATDEP5')).enabled := bGrise;
   THLabel(GetControl('TYMT_CATLOC1')).enabled := bGrise;
   THLabel(GetControl('TYMT_CATLOC2')).enabled := bGrise;
   THLabel(GetControl('TYMT_CATLOC3')).enabled := bGrise;
   THLabel(GetControl('TYMT_CATLOC4')).enabled := bGrise;
   THLabel(GetControl('TYMT_CATLOC5')).enabled := bGrise;
   THEdit (GetControl('YMT_LIBCAT1')).enabled := bGrise ;
   THEdit (GetControl('YMT_LIBCAT2')).enabled := bGrise ;
   THEdit (GetControl('YMT_LIBCAT3')).enabled := bGrise ;
   THEdit (GetControl('YMT_LIBCAT4')).enabled := bGrise ;
   THEdit (GetControl('YMT_LIBCAT5')).enabled := bGrise ;
   ZYMT_CRI1LOC.Enabled := bGrise;
   ZYMT_CRI2LOC.Enabled := bGrise;
   ZYMT_CRI3LOC.Enabled := bGrise;
   ZYMT_CRI4LOC.Enabled := bGrise;
   ZYMT_CRI5LOC.Enabled := bGrise;
   ZYMT_CAT1DEP.Enabled := bGrise;
   ZYMT_CAT2DEP.Enabled := bGrise;
   ZYMT_CAT3DEP.Enabled := bGrise;
   ZYMT_CAT4DEP.Enabled := bGrise;
   ZYMT_CAT5DEP.Enabled := bGrise;
{$ENDIF}
end;


Initialization
  registerclasses ( [ TOM_YMODELETAXE ] ) ;

{$ENDIF}

end.

