{ Unité : Source TOM de la TABLE : CPTYPEVISA
--------------------------------------------------------------------------------------
    Version   |  Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
 7.01.001.001  02/02/06  JP   Création de l'unité
 8.01.001.001  07/12/06  JP   FQ 19162 : Au rechargement d'un enregistrement, les champs CTI_CODELIBRE sont donc vidés
 8.01.001.018  30/05/07  JP   L'établissement est maintenant facultatif
 8.01.001.019  15/06/07  JP   Filtre sur les circuit lors du changement du nombre de visas
 8.01.001.021  20/06/07  JP   FQ 20778 : Les natures de pièces ne sont plus obligatoires
 8.01.001.021  21/06/07  JP   FQ 20777 : Incrémentation automatique du code
--------------------------------------------------------------------------------------}
unit CPTYPEVISA_TOM;

interface

uses
  StdCtrls, Controls, Classes,
  {$IFDEF EAGLCLIENT}
  MaineAGL, eFiche,
  {$ELSE}
  FE_Main, Fiche, db, HDB, {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
  {$ENDIF}
  {$IFDEF VER150}
  Variants,
  {$ENDIF}
  UObjGen, 
  SysUtils, UTOM, Grids, HCtrls, UTob;

const
  TypChar = ['0'..'9', 'A'..'Z', ':', '-', ','];

type
  TOM_CPTYPEVISA = class(TOM)
    FListe   : THGrid; {Pour la grille, s'inspirer de ce qui se fait dans rubrique_tom}
    {$IFDEF EAGLCLIENT}
    TypeLibre1  : THValComboBox;
    TypeLibre2  : THValComboBox;
    TypeLibre3  : THValComboBox;
    CodeLibre1  : THValComboBox;
    CodeLibre2  : THValComboBox;
    CodeLibre3  : THValComboBox;
    AxeLibre1   : THValComboBox;
    AxeLibre2   : THValComboBox;
    AxeLibre3   : THValComboBox;
    TexteLibre1 : THEdit;
    TexteLibre2 : THEdit;
    TexteLibre3 : THEdit;
    {$ELSE}
    TypeLibre1  : THDBValComboBox;
    TypeLibre2  : THDBValComboBox;
    TypeLibre3  : THDBValComboBox;
    CodeLibre1  : THDBValComboBox;
    CodeLibre2  : THDBValComboBox;
    CodeLibre3  : THDBValComboBox;
    AxeLibre1   : THDBValComboBox;
    AxeLibre2   : THDBValComboBox;
    AxeLibre3   : THDBValComboBox;
    TexteLibre1 : THDBEdit;
    TexteLibre2 : THDBEdit;
    TexteLibre3 : THDBEdit;
    {$ENDIF EAGLCLIENT}

    procedure OnNewRecord              ; override;
    procedure OnDeleteRecord           ; override;
    procedure OnUpdateRecord           ; override;
    procedure OnAfterUpdateRecord      ; override;
    procedure OnLoadRecord             ; override;
    procedure OnChangeField(F : TField); override;
    procedure OnArgument   (S : string); override;
    procedure OnCancelRecord           ; override;
    procedure OnClose                  ; override;
  private
    ValeurEnCours : string;
    ObjetCode     : TObjCodeCombo; {21/06/07 : FQ 20777}

    function  Grid2St(ACol : Integer) : string;
    function  TestGrille              : Boolean;
    function  ControlZonesLibres      : Boolean;

    procedure MajAffichage(Value, Ind : string; CodeSender : Boolean = False);
    procedure RecupControls;
    procedure RecupEvenements;
    procedure TronqueLeCompte(var Chaine : string);
    procedure EcritCompte;
    procedure LitCompte;
    procedure St2Grid(St : string; Col : Byte);
    procedure MetEnModification;
  public
    procedure FlisteOnKeyPress(Sender : TObject; var Key: Char);
    procedure FlisteOnKeyDown (Sender : TObject; var Key: Word; Shift: TShiftState);
    procedure ZonesOnEnter    (Sender : TObject);
    procedure ZonesOnExit     (Sender : TObject);
    procedure NbVisasChange   (NewNb : Integer); {15/06/07}
  end;

procedure CpLanceFiche_TypeVisaFiche(Range, Lequel, Argument : string);


implementation

uses
  HEnt1, HMsgBox, HTB97, Ent1, uLibBonAPayer, Windows, Forms;

{---------------------------------------------------------------------------------------}
procedure CpLanceFiche_TypeVisaFiche(Range, Lequel, Argument : string);
{---------------------------------------------------------------------------------------}
begin
  AGLLanceFiche('CP', 'CPTYPEVISA', Range, Lequel, Argument);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPTYPEVISA.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 7509015;

  RecupControls;
  RecupEvenements;
  
  if TFFiche(Ecran).FTypeAction = taConsult then
    FListe.Options := FListe.Options - [goTabs, goEditing] + [goRowSelect];
  TFFiche(Ecran).Pages.ActivePageIndex := 0;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPTYPEVISA.OnCancelRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  {Rafraîchit la grille}
  LitCompte;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPTYPEVISA.OnChangeField(F : TField);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if (UpperCase(F.FieldName) = 'CTI_TYPELIBRE1') or
     (UpperCase(F.FieldName) = 'CTI_TYPELIBRE2') or
     (UpperCase(F.FieldName) = 'CTI_TYPELIBRE3') then
    MajAffichage(F.AsString, F.FieldName[Length(F.FieldName)])
    
  else if (UpperCase(F.FieldName) = 'CTI_CODELIBRE1') or
          (UpperCase(F.FieldName) = 'CTI_CODELIBRE2') or
          (UpperCase(F.FieldName) = 'CTI_CODELIBRE3') then
    MajAffichage(GetField('CTI_TYPELIBRE' + F.FieldName[Length(F.FieldName)]), F.FieldName[Length(F.FieldName)], True)

  else if UpperCase(F.FieldName) = 'CTI_NBVISA' then {15/06/07}
    NbVisasChange(F.AsInteger);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPTYPEVISA.OnDeleteRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if not CanDeleteTypeVisa(VarToStr(GetField('CTI_CODEVISA'))) then begin
    LastError := 2;
    LastErrorMsg := TraduireMemoire('Ce type de visa est utilisé dans les bons à payer et ne peut être supprimé.');
  end
  else
    FListe.VidePile(True);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPTYPEVISA.OnLoadRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  {Charge la grille}
  LitCompte;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPTYPEVISA.OnClose;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(ObjetCode) then FreeAndNil(ObjetCode);
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPTYPEVISA.OnNewRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  SetField('CTI_CIRCUITBAP'   , THValComboBox(GetControl('CTI_CIRCUITBAP'   )).Values[0]);
  SetField('CTI_ETABLISSEMENT', THValComboBox(GetControl('CTI_ETABLISSEMENT')).Values[0]);
  SetField('CTI_NATUREPIECE'  , THValComboBox(GetControl('CTI_NATUREPIECE'  )).Values[0]);
  SetField('CTI_NBVISA', 1);
  SetField('CTI_AUCHOIX', 'X');

  {21/06/07 : FQ 20777 : génération automatique du Code}
  if not Assigned(ObjetCode) then begin
    ObjetCode := TObjCodeCombo.Create('CPTYPEVISA', 'CTI_CODEVISA');
    ObjetCode.LastCode := '';
  end;
  SetField('CTI_CODEVISA', ObjetCode.GetNewCode);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPTYPEVISA.OnAfterUpdateRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  AvertirTable('CPTYPEVISABAP');
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPTYPEVISA.OnUpdateRecord;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
  D : Double;
begin
  inherited;
  {Contrôles simples sur les champs obligatoires}
  if VarToStr(GetField('CTI_CIRCUITBAP')) = '' then begin
    LastError := 3;
    LastErrorMsg := TraduireMemoire('Veuillez choisir un circuit.');
    SetFocusControl('CTI_CIRCUITBAP');
  end
  {30/05/07 : l'établissement est facultatif à la demande de SIC
  else if VarToStr(GetField('CTI_ETABLISSEMENT')) = '' then begin
    LastError := 3;
    LastErrorMsg := TraduireMemoire('Veuillez choisir un établissement.');
    SetFocusControl('CTI_ETABLISSEMENT');
  end
  20/06/07 : FQ 20778 : Nouvelle demande de SIC pour les natures de pièce
  else if VarToStr(GetField('CTI_NATUREPIECE')) = '' then begin
    LastError := 3;
    LastErrorMsg := TraduireMemoire('Veuillez choisir une nature de pièce.');
    SetFocusControl('CTI_NATUREPIECE');
  end}
  else if VarToStr(GetField('CTI_LIBELLE')) = '' then begin
    LastError := 3;
    LastErrorMsg := TraduireMemoire('Veuillez saisir un libellé.');
    SetFocusControl('CTI_LIBELLE');
  end
  else if ValeurI(GetField('CTI_NBVISA')) <= 0 then begin
    LastError := 3;
    LastErrorMsg := TraduireMemoire('Le nombre de visas doit être supérieur à 0.');
    SetFocusControl('CTI_NBVISA');
  end
  else if Valeur(GetField('CTI_MONTANTMIN')) > Valeur(GetField('CTI_MONTANTMAX')) then begin
    D := Valeur(GetField('CTI_MONTANTMIN'));
    SetField('CTI_MONTANTMIN', Valeur(GetField('CTI_MONTANTMAX')));
    SetField('CTI_MONTANTMAX', D);
  end;

  if LastError <> 0 then Exit;

  {Ecriture du contenu de la grille dans les Champs COMPTE et EXCLUSION}
  if TestGrille then
    EcritCompte
  else begin
    LastError := 3;
    LastErrorMsg := TraduireMemoire('La saisie dans la grille n''est pas correcte.'#13 +
                    'Vous ne pouvez saisir que des chiffres, des lettres et les caractère "-" , ":" et ","');
    SetFocusControl('FLISTE');
  end;

  if LastError <> 0 then Exit;

  {Contrôle de la saisie sur les zones libres}
  if not ControlZonesLibres then begin
    LastError := 3;
    LastErrorMsg := TraduireMemoire('La saisie des zones libres n''est pas correcte.');
  end;

  if LastError <> 0 then Exit;

  {On s'assure que le nombre de visas correspond à celui du circuit}
  Q := OpenSQL('SELECT COUNT(*) FROM CPCIRCUIT WHERE CCI_CIRCUITBAP = "' + VarToStr(GetField('CTI_CIRCUITBAP')) + '"', True);
  try
    if Q.Fields[0].AsInteger <> ValeurI(GetField('CTI_NBVISA')) then begin
      LastError := 3;
      LastErrorMsg := TraduireMemoire('Le nombre de visas (' + VarToStr(GetField('CTI_NBVISA')) + ') n''est pas cohérent'#13 +
                                      'avec le nombre de visas du cricuit (' + Q.Fields[0].AsString + ').' );
      SetFocusControl('CTI_NBVISA');
    end;
  finally
    Ferme(Q);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPTYPEVISA.FlisteOnKeyPress(Sender : TObject; var Key : Char);
{---------------------------------------------------------------------------------------}
begin
  if TFFiche(Ecran).FTypeAction = taConsult then Exit;
  if not (Key in TypChar) and (Key <> #8 {Back}) then
    Key := #0
  else begin
    Key := UpCase(Key);
    MetEnModification;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPTYPEVISA.FlisteOnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
{---------------------------------------------------------------------------------------}
begin
  if TFFiche(Ecran).FTypeAction = taConsult then Exit;

  case Key of
    VK_DOWN   : begin
                  if (Fliste.Row = FListe.RowCount - 1) then begin
                    MetEnModification;
                    FListe.InsertRow(Fliste.Row + 1);
                  end;
                end;
    VK_INSERT : begin
                  MetEnModification;
                  FListe.InsertRow(Fliste.Row + 1);
                end;
    VK_DELETE : begin
                  MetEnModification;
                  FListe.DeleteRow(Fliste.Row);
                end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPTYPEVISA.ZonesOnEnter(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  ValeurEnCours := GetField(TWinControl(Sender).Name);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPTYPEVISA.ZonesOnExit(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  c : Char;
begin
  {Si la valeur du champ a changé, remise à vides des champs dépendants}
  if ValeurEnCours <> GetField(TWinControl(Sender).Name) then begin
    c := TWinControl(Sender).Name[Length(TWinControl(Sender).Name)];
    {Zone tablette sur les tables libres des sections}
    if Pos('AXE', UpperCase(TWinControl(Sender).Name)) > 0 then 
      SetField('CTI_TEXTELIBRE' + c, '')
    {Zone tablette sur les tables libres}
    else if Pos('CODE', UpperCase(TWinControl(Sender).Name)) > 0 then begin
      SetField('CTI_TEXTELIBRE' + c, '');
      SetField('CTI_AXE'        + c, '');
    end
    {Zone tablette sur les types de zones libres}
    else begin
      SetField('CTI_TEXTELIBRE' + c, '');
      SetField('CTI_CODELIBRE'  + c, '');
      SetField('CTI_AXE'        + c, '');
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
function TOM_CPTYPEVISA.Grid2St(ACol : Integer) : string;
{---------------------------------------------------------------------------------------}
var
  St    : string;
  StFin : string;
  ARow  : Integer;
begin
  if ACol >= FListe.ColCount then begin
    Result := StFin;
    Exit;
  end;

  for ARow := 1 to FListe.RowCount - 1  do begin
    {Si la ligne est vide, on passe à la suivante}
    if (FListe.Cells[0, ARow] = '') and (FListe.Cells[1, ARow] = '') then
      Continue;
   {S'il n'y a pas de compte de sélection, mais des comptes d'exclusion}
   if (FListe.Cells[0, ARow] = '') and (FListe.Cells[1, ARow] <> '')then begin
     HShowMessage('0;' + Ecran.Caption + ';La ligne ' + IntToStr(ARow) + ' n''a pas de compte de sélection'#13 +
                  'mais des comptes d''exclusion : elle sera ignorée;I;O;O;O;', '', '');
     Continue;
   end;

    St := FListe.Cells[ACol, ARow];
    if St <> '' then StFin := StFin + St;
    if (StFin <> '') or (FListe.RowCount > 1) then StFin := StFin + ';'; {JP 24/01/07 : FQ 19444}
  end;

  if Length(StFin) > 250 then
    TronqueLeCompte(StFin);
  Result := StFin;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPTYPEVISA.EcritCompte;
{---------------------------------------------------------------------------------------}
begin
  SetField('CTI_COMPTE'   , Grid2St(0));
  SetField('CTI_EXCLUSION', Grid2St(1));
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPTYPEVISA.LitCompte;
{---------------------------------------------------------------------------------------}
begin
  Fliste.VidePile(True);
  FListe.RowCount := 2;
  St2Grid(GetField('CTI_COMPTE')   , 0);
  St2Grid(GetField('CTI_EXCLUSION'), 1);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPTYPEVISA.St2Grid(St : string; Col : Byte);
{---------------------------------------------------------------------------------------}
var
  St1 : string;
  n   : Integer;
begin
  n := 1;
  while St <> '' do begin
    St1 := Trim(ReadTokenSt(St));

    FListe.Cells[Col, n] := St1;
    Inc(n);
    if n > FListe.RowCount - 1 then
      FListe.RowCount := FListe.RowCount + 1;
  end;
end;

{---------------------------------------------------------------------------------------}
function TOM_CPTYPEVISA.TestGrille : Boolean;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  p : Integer;
  k : Integer;
  s : string;
begin
  Result := True;
  for p := 0 to 1 do
    for n := 1 to FListe.RowCount - 1 do begin
      s := Trim(FListe.Cells[p, n]);
      k := 1;
      while k <= Length(s)  do begin
        if not (s[k] in TypChar) then begin
          Result := False;
          Exit;
        end;
        Inc(k);
      end;
    end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPTYPEVISA.TronqueLeCompte(var Chaine : string);
{---------------------------------------------------------------------------------------}
var
  St,
  St1 : string;
begin
  while Chaine <> '' do begin
    St1 := ReadTokenSt(Chaine);
    if Length(St) + Length(St1) + 1 > 249 then Break;
    St := St + St1 + ';';
  end;
  Chaine := St;
end;

{---------------------------------------------------------------------------------------}
function TOM_CPTYPEVISA.ControlZonesLibres : Boolean;
{---------------------------------------------------------------------------------------}

   {-------------------------------------------------------------------------}
   function TesteZonesInduites(Value, Ind : string) : Boolean;
   {-------------------------------------------------------------------------}
   begin
     Result := True;
     if Value <> '' then begin
       {S'il y a une valeur à type libre, on s'assure que les 2 champs dépendants sont renseignés}
       if (Value = tyt_Devise) or (Value = tyt_RefInterne) then begin
         Result := Trim(VarToStr(GetField('CTI_TEXTELIBRE' + Ind))) <> '';
         {Et on vide eventuellement le code et l'axe qui n'a pas à être rempli}
         SetField('CTI_CODELIBRE' + Ind, '');
         SetField('CTI_AXE'       + Ind, '');
       end
       else if Value = tyt_TLSection then begin
         Result := (Trim(VarToStr(GetField('CTI_TEXTELIBRE' + Ind))) <> '') and
                   (Trim(VarToStr(GetField('CTI_AXE'        + Ind))) <> '') and
                   (Trim(VarToStr(GetField('CTI_CODELIBRE'  + Ind))) <> '');

       end
       else begin
         Result := (Trim(VarToStr(GetField('CTI_TEXTELIBRE' + Ind))) <> '') and
                   (Trim(VarToStr(GetField('CTI_CODELIBRE'  + Ind))) <> '');
       end;

       if not Result then SetFocusControl('CTI_TEXTELIBRE' + Ind);
     end
     else begin
       {S'il n'y a pas de valeur à type libre, on force à vide les 2 champs dépendants}
       SetField('CTI_CODELIBRE'  + Ind, '');
       SetField('CTI_TEXTELIBRE' + Ind, '');
       SetField('CTI_AXE'        + Ind, '');
     end;
   end;

begin
  Result := True;
  Result := Result and TesteZonesInduites(VarToStr(GetField('CTI_TYPELIBRE1')), '1')
                   and TesteZonesInduites(VarToStr(GetField('CTI_TYPELIBRE2')), '2')
                   and TesteZonesInduites(VarToStr(GetField('CTI_TYPELIBRE3')), '3')
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPTYPEVISA.MajAffichage(Value, Ind : string; CodeSender : Boolean = False);
{---------------------------------------------------------------------------------------}
var
  Code   : string;
  AxeOk  : Boolean;
begin
  if Value = '' then begin
    SetControlVisible('CTI_CODELIBRE'   + Ind, True);
    SetControlVisible('TCTI_CODELIBRE'  + Ind, True);
    SetControlVisible('CTI_AXE'         + Ind, True);
    SetControlVisible('TCTI_AXE'        + Ind, True);
    SetControlEnabled('CTI_AXE'         + Ind, False);
    SetControlEnabled('TCTI_AXE'        + Ind, False);
    SetControlEnabled('CTI_CODELIBRE'   + Ind, False);
    SetControlEnabled('TCTI_CODELIBRE'  + Ind, False);
    SetControlEnabled('TCTI_TEXTELIBRE' + Ind, False);
    SetControlEnabled('CTI_TEXTELIBRE'  + Ind, False);
    SetControlProperty('TCTI_TEXTELIBRE' + Ind, 'LEFT', 324);
    SetControlProperty('CTI_TEXTELIBRE' + Ind, 'LEFT', 324);
    SetControlProperty('CTI_TEXTELIBRE' + Ind, 'WIDTH', 131);
    SetControlProperty('CTI_TEXTELIBRE' + Ind, 'ELIPSISBUTTON', False);
    SetControlProperty('CTI_TEXTELIBRE' + Ind, 'DATATYPE', '');
  end

  else begin
    SetControlEnabled('TCTI_TEXTELIBRE' + Ind, True);
    SetControlEnabled('CTI_TEXTELIBRE'  + Ind, True);

    if (Value = tyt_Devise) or (Value = tyt_RefInterne) then begin
      if TFFiche(Ecran).Pages.ActivePageIndex = 1 then
        SetFocusControl('CTI_TEXTELIBRE'   + Ind);
      SetControlVisible('CTI_CODELIBRE'    + Ind, False);
      SetControlVisible('TCTI_CODELIBRE'   + Ind, False);
      SetControlVisible('TCTI_AXE'         + Ind, False);
      SetControlVisible('CTI_AXE'          + Ind, False);
      SetControlProperty('TCTI_TEXTELIBRE' + Ind, 'LEFT', 112);
      SetControlProperty('CTI_TEXTELIBRE'  + Ind, 'LEFT', 112);
      SetControlProperty('CTI_TEXTELIBRE'  + Ind, 'WIDTH', 343);
      SetControlProperty('CTI_TEXTELIBRE'  + Ind, 'ELIPSISBUTTON', False);
      SetControlProperty('CTI_TEXTELIBRE'  + Ind, 'DATATYPE', '');
    end

    else begin
      SetControlProperty('TCTI_TEXTELIBRE' + Ind, 'LEFT', 219);
      SetControlProperty('CTI_TEXTELIBRE'  + Ind, 'LEFT', 219);
      SetControlProperty('CTI_TEXTELIBRE'  + Ind, 'WIDTH', 236);
      SetControlProperty('CTI_TEXTELIBRE'  + Ind, 'ELIPSISBUTTON', True);

      AxeOk := True;
      {Récupération éventuelle de l'indice de la table libre à traiter}
      {08/12/06 : FQ 19162 : en eAGL, au chargement de l'enregistrement, GetControlText renvoie vide => GetField}
      Code := GetField('CTI_CODELIBRE' + Ind);
      if Value = tyt_TLSection then begin
        AxeOk := GetControlVisible('CTI_AXE' + Ind);
        SetControlVisible('CTI_AXE'  + Ind, True);
        SetControlVisible('TCTI_AXE' + Ind, True);
        SetControlEnabled('CTI_AXE'  + Ind, True);
        SetControlEnabled('TCTI_AXE' + Ind, True);
      end
      else begin
        SetControlVisible('TCTI_AXE' + Ind, False);
        SetControlVisible('CTI_AXE'  + Ind, False);
      end;

      if not CodeSender then begin
        {Le sender est le type, on force l'ActiveControl sur le code}
        if not GetControlVisible('CTI_CODELIBRE' + Ind) then begin
          SetControlVisible('CTI_CODELIBRE'  + Ind, True);
          SetControlVisible('TCTI_CODELIBRE' + Ind, True);
          if TFFiche(Ecran).Pages.ActivePageIndex = 1 then
            SetFocusControl('CTI_CODELIBRE' + Ind);
        end;

        if not GetControlEnabled('CTI_CODELIBRE' + Ind) then begin
          SetControlEnabled('CTI_CODELIBRE'   + Ind, True);
          SetControlEnabled('TCTI_CODELIBRE'  + Ind, True);
          if TFFiche(Ecran).Pages.ActivePageIndex = 1 then
            SetFocusControl('CTI_CODELIBRE' + Ind);
        end;
      end
      else if Value = tyt_TLSection then begin
        {Pour ne pas faire le SetFocus si on ne passe pas au contrôle suivant, Maj + Tab par exemple}
        if (TFFiche(Ecran).Pages.ActivePageIndex = 1) and not AxeOk then
          SetFocusControl('CTI_AXE' + Ind);
      end;

      if Length(Code) = 3  then Code := Code[3]
                           else Code := '0';

      if Value = tyt_TLAuxiliaire then begin
        SetControlProperty('CTI_CODELIBRE'  + Ind, 'DATATYPE', 'TTTABLESLIBRESAUX');
        SetControlProperty('CTI_TEXTELIBRE' + Ind, 'DATATYPE', 'TZNATTIERS' + Code);
      end
      else if Value = tyt_TLEcriture then begin
        SetControlProperty('CTI_CODELIBRE'  + Ind, 'DATATYPE', 'TTTABLESLIBRESECR');
        SetControlProperty('CTI_TEXTELIBRE' + Ind, 'DATATYPE', 'TZNATECR' + Code);
      end
      else if Value = tyt_TLGeneral then begin
        SetControlProperty('CTI_CODELIBRE'  + Ind, 'DATATYPE', 'TTTABLESLIBRESGEN');
        SetControlProperty('CTI_TEXTELIBRE' + Ind, 'DATATYPE', 'TZNATGENE' + Code);
      end
      else if Value = tyt_TLSection    then begin
        SetControlProperty('CTI_CODELIBRE' + Ind, 'DATATYPE', 'TTTABLESLIBRESSEC');
        SetControlProperty('TCTI_TEXTELIBRE' + Ind, 'LEFT', 324);
        SetControlProperty('CTI_TEXTELIBRE'  + Ind, 'LEFT', 324);
        SetControlProperty('CTI_TEXTELIBRE'  + Ind, 'WIDTH', 131);
        SetControlProperty('CTI_TEXTELIBRE'  + Ind, 'DATATYPE', 'TZNATSECT' + Code);
      end;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPTYPEVISA.RecupControls;
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF EAGLCLIENT}
  TypeLibre1  := THValComboBox(GetControl('CTI_TYPELIBRE1'));
  TypeLibre2  := THValComboBox(GetControl('CTI_TYPELIBRE2'));
  TypeLibre3  := THValComboBox(GetControl('CTI_TYPELIBRE3'));
  CodeLibre1  := THValComboBox(GetControl('CTI_CODELIBRE1'));
  CodeLibre2  := THValComboBox(GetControl('CTI_CODELIBRE2'));
  CodeLibre3  := THValComboBox(GetControl('CTI_CODELIBRE3'));
  AxeLibre1   := THValComboBox(GetControl('CTI_AXE1'));
  AxeLibre2   := THValComboBox(GetControl('CTI_AXE2'));
  AxeLibre3   := THValComboBox(GetControl('CTI_AXE3'));
  TexteLibre1 := THEdit(GetControl('CTI_TEXTELIBRE1'));
  TexteLibre2 := THEdit(GetControl('CTI_TEXTELIBRE2'));
  TexteLibre3 := THEdit(GetControl('CTI_TEXTELIBRE3'));
  {$ELSE}
  TypeLibre1  := THDBValComboBox(GetControl('CTI_TYPELIBRE1'));
  TypeLibre2  := THDBValComboBox(GetControl('CTI_TYPELIBRE2'));
  TypeLibre3  := THDBValComboBox(GetControl('CTI_TYPELIBRE3'));
  CodeLibre1  := THDBValComboBox(GetControl('CTI_CODELIBRE1'));
  CodeLibre2  := THDBValComboBox(GetControl('CTI_CODELIBRE2'));
  CodeLibre3  := THDBValComboBox(GetControl('CTI_CODELIBRE3'));
  AxeLibre1   := THDBValComboBox(GetControl('CTI_AXE1'));
  AxeLibre2   := THDBValComboBox(GetControl('CTI_AXE2'));
  AxeLibre3   := THDBValComboBox(GetControl('CTI_AXE3'));
  TexteLibre1 := THDBEdit(GetControl('CTI_TEXTELIBRE1'));
  TexteLibre2 := THDBEdit(GetControl('CTI_TEXTELIBRE2'));
  TexteLibre3 := THDBEdit(GetControl('CTI_TEXTELIBRE3'));
  {$ENDIF EAGLCLIENT}
  FListe := THGrid(GetControl('FLISTE'));
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPTYPEVISA.RecupEvenements;
{---------------------------------------------------------------------------------------}
begin
  TypeLibre1.OnEnter := ZonesOnEnter;
  TypeLibre2.OnEnter := ZonesOnEnter;
  TypeLibre3.OnEnter := ZonesOnEnter;
  CodeLibre1.OnEnter := ZonesOnEnter;
  CodeLibre2.OnEnter := ZonesOnEnter;
  CodeLibre3.OnEnter := ZonesOnEnter;
  AxeLibre1.OnEnter  := ZonesOnEnter;
  AxeLibre2.OnEnter  := ZonesOnEnter;
  AxeLibre3.OnEnter  := ZonesOnEnter;

  TypeLibre1.OnExit := ZonesOnExit;
  TypeLibre2.OnExit := ZonesOnExit;
  TypeLibre3.OnExit := ZonesOnExit;
  CodeLibre1.OnExit := ZonesOnExit;
  CodeLibre2.OnExit := ZonesOnExit;
  CodeLibre3.OnExit := ZonesOnExit;
  AxeLibre1.OnExit  := ZonesOnExit;
  AxeLibre2.OnExit  := ZonesOnExit;
  AxeLibre3.OnExit  := ZonesOnExit;

  FListe.OnKeyDown  := FlisteOnKeyDown;
  Fliste.OnKeyPress := FlisteOnKeyPress;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CPTYPEVISA.MetEnModification;
{---------------------------------------------------------------------------------------}
begin
  if not (DS.State in [dsEdit, dsInsert]) then DS.Edit;
end;

{15/06/07 : On filtre les circuits pour n'afficher que ceux qui ont le nombre d'étapes demandées
{---------------------------------------------------------------------------------------}
procedure TOM_CPTYPEVISA.NbVisasChange(NewNb : Integer);
{---------------------------------------------------------------------------------------}
var
  {$IFDEF EAGLCLIENT}
  CB : THValComboBox;
  {$ELSE}
  CB : THDBValComboBox;
  {$ENDIF EAGLCLIENT}
begin
  {$IFDEF EAGLCLIENT}
  CB := (GetControl('CTI_CIRCUITBAP') as THValComboBox);
  {$ELSE}
  CB := (GetControl('CTI_CIRCUITBAP') as THDBValComboBox);
  {$ENDIF EAGLCLIENT}

  if NewNb <= 0 then
    CB.Plus := ''
  else
    CB.Plus := 'AND CCI_CIRCUITBAP IN (SELECT CCI_CIRCUITBAP FROM CPVUECIRCUIT ' +
               'WHERE CCI_NBVISAS = ' + IntToStr(NewNb) + ') ';

  if DS.State in [dsEdit, dsInsert] then begin
    CB.ItemIndex := 0;
    SetField('CTI_CIRCUITBAP', CB.Value);
  end;
end;

initialization
  RegisterClasses([TOM_CPTYPEVISA]);

end.
