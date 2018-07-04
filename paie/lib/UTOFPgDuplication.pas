{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Unit générique de duplication du paramètrage de tous les
Suite ........ : paramètrages de la paie.
Suite ........ : Spécifique à la paie car prend en compte que les tables de
Suite ........ : la paie et controle par rapport à dos,std,ceg les codes
Suite ........ : numériques à dupliquer en tenant compte des codes pairs
Suite ........ : ou impairs..;
Suite ........ : Au moins, il y a une normalisation de la fiche de duplication
Suite ........ : et des méthodes
Mots clefs ... : PAIE;DUPLICATION
*****************************************************************}
{
PT1 : 01/10/2001 : V562  MF  Ajout du traitement de duplication des
                             codifications de la Ducs
PT2 : 18/12/2001 : V569  MF  Duplication des codoif DUCS (Pas de StrToInt donc
                             pas d'appel à DefautCode)
PT3 : 09/10/2002 : V585  PH  Duplication des elts nationaux autorisée avec le type DOS
PT4   20/02/2003   V_42C SB  FQ 10526 l'utilisateur non Cegid ni réviseur ne duplique qu'en DOS
PT5   02/08/2004   V_60  PH  FQ 12487 Ergonomie
PT6   03/10/2005   V_60  PH  FQ 12623 Controle du code profil en duplication si alpha
PT7   11/10/2005   V_60  PH  FQ 12623 Si alpha et profil uniquement STD.
PT8   20/10/2006   V_70  SB  FQ 13390 Impossibilite de dupliquer en alpha les cot et les rem
PT9   06/03/2007   V_70  FC  Rajout de la convention collective pour les éléments nationaux
PT10  12/06/2007   V_72  FC  Gestion du niveau préconisé pour filter la combo Predefini
PT11  04/07/2007   V_72  FC  FQ 14475 Revue des concepts
PT12  17/07/2007   V_8   GGU Gestion de la duplication des variables de présence
PT13  03/09/2007   V_80  FC  FQ 14721 Gestion de la convention collective dans le prédéfini CEG
PT14  04/09/2007   V_80  FC  FQ 14605 Duplication d'un établissement
PT15  20/11/2007   V_80  FC  FQ 14956 Duplication élément national, initialiser la convention collective
PT16  26/05/2008   V_850 FC  FQ 15354 Duplication d'un organisme
}

unit UTOFPgDuplication;

interface
uses Classes, sysutils, HTB97, HCtrls, HMsgBox, UTOF, ParamDat, Spin
{$IFNDEF EAGLCLIENT}
{$IFNDEF DBXPRESS} ,dbtables {$ELSE} ,uDbxDataSet {$ENDIF}
{$ELSE}

{$ENDIF}
,Utob
;

type
  TOF_PGDUPLICATION = class(TOF)
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnClose; override;
    procedure Validation(Sender: TObject);
    procedure DateElipsisclick(Sender: TObject);
    procedure DefautCode(Sender: TObject);
    function ExceptionEltNat(code: string): boolean;
    procedure AffectCode(Sender: TObject);
    procedure PredefiniChange(Sender : TObject);  //PT13
  private
    CEG, STD, DOS: Boolean;
    NomRubrique, AncCode, AncPred: string;
    MaxRub, Longueur: integer;
    procedure MajFiltreNiveau(CodeElt : String); //PT10
  end;

//DEB PT14
type
  TOF_PGDUPLIETAB = class(TOF)
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnClose; override;
    procedure Validation(Sender: TObject);
  end;
//FIN PT14

//DEB PT16
type
  TOF_PGDUPLIORG = class(TOF)
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnClose; override;
    procedure Validation(Sender: TObject);
  private
    TypeDupli,EtabOri : String;
  end;
//FIN PT16

implementation

uses Pgoutils, Hent1, PgOutils2, StrUtils;

{ TOF_PGDUPLICATION }





procedure TOF_PGDUPLICATION.DateElipsisclick(Sender: TObject);
var key: char;
begin
  key := '*';
  ParamDate(Ecran, Sender, Key);
end;

procedure TOF_PGDUPLICATION.DefautCode(Sender: TObject);
var
  nb: integer;
begin
  if Length(GetControlText('CODEDUPLIQUER')) < Longueur then
  begin
    nb := Longueur - Length(GetControlText('CODEDUPLIQUER'));
    SetControlText('CODEDUPLIQUER', UpperCase(StringOfChar('0', nb) + GetControlText('CODEDUPLIQUER')));
  end;
{Ecode := THEdit(GetControl('CODEDUPLIQUER'));
If (ECode<>nil) then
  if Ecode.text<>'' then
    SetControlProperty('CODEDUPLIQUER','Text',ColleZeroDevant(StrToInt(Ecode.text),Longueur));}
end;

procedure TOF_PGDUPLICATION.AffectCode(Sender: TObject);
begin
  SetControlText('CODEDUPLIQUER', GetControlText('CODESPIN'));
end;

function TOF_PGDUPLICATION.ExceptionEltNat(code: string): boolean;
var icode: integer;
begin
  icode := StrtoInt(code);
  if (icode >= 300) and (icode <= 315) then result := true else result := False;
end;

procedure TOF_PGDUPLICATION.OnArgument(Arguments: string);
var
  ECode, EDate: THEdit;
  Btn: TToolBarButton97;
  CodeSpin: TSpinEdit;
  THPredefini : THValComboBox;
  ConvCol:String;//PT15
begin
  DonneCodeDupliquer('', '', '',''); //PT9
  Longueur := 0;
  NomRubrique := ReadTokenst(Arguments);
  AncCode := ReadTokenst(Arguments);
  AncPred := ReadTokenst(Arguments);
  //DEB PT15
  if NomRubrique = 'ELT' then
  begin
    Longueur := StrToInt(ReadTokenst(Arguments));
    ConvCol := Arguments;
  end
  //FIN PT15
  else
  Longueur := StrToInt(Arguments);
  MaxRub := 0;

  if (Longueur = 3) and (NomRubrique <> 'BILAN') then MaxRub := 100; //3 pour le nd de chiffre des profils
  if (Longueur = 4) then MaxRub := 1000; //elt nationaux
  if (NomRubrique = 'COT') or (NomRubrique = 'AAA') or (NomRubrique = 'VENT') then Maxrub := 0; ////Cotisation,rémunération,
  if Longueur = 0 then SetControlProperty('CODEDUPLIQUER', 'MaxLength', 17) //17 pour varchar(17)
  else SetControlProperty('CODEDUPLIQUER', 'MaxLength', Longueur);
  SetControlEnabled('PREDEFINI', True);
  SetControlProperty('PREDEFINI', 'value', 'DOS');
  SetControlProperty('DATEVALIDITE', 'text', Date);
  if (NomRubrique = 'VAR') and (AncCode[1] = 'P') then  //PT12 Cas des variables de présence
  begin
    SetControlProperty('CODEDUPLIQUER', 'EditMask', 'P000;1;_');
  end;
  AccesPredefini('TOUS', CEG, STD, DOS);
  if NomRubrique = 'ELT' then
  begin
    //DEB PT9
    SetControlVisible('TCONVENTION', True);
    SetControlVisible('CONVENTION', True);
    //FIN PT9
    SetControlVisible('LBDATEVALIDITE', True);
    SetControlVisible('DATEVALIDITE', True);
    SetControlEnaBled('CODEDUPLIQUER', False);
    SetControlProperty('CODEDUPLIQUER', 'TEXT', AncCode);
// PT3 : 09/10/2002 : V585  PH  Duplication des elts nationaux autorisée avec le type DOS
    SetControlEnabled('PREDEFINI', TRUE); // FALSE
    SetControlProperty('PREDEFINI', 'DataType', 'YYPREDEFINI'); // YYPREDCEGSTD
// FIN PT3
    SetControlProperty('PREDEFINI', 'value', AncPred);
    if (CEG = False) and (STD = False) then
    begin { PT4 Non accès au utilisateur non Cegid ni Réviseur }
      SetControlProperty('PREDEFINI', 'value', 'DOS');
      SetControlEnabled('PREDEFINI', FALSE);
    end;
  end
  //DEB PT9
  else
  begin
    SetControlVisible('TCONVENTION', False);
    SetControlVisible('CONVENTION', False);
    //FIN PT9
  end;
  if NomRubrique = 'VENT' then
  begin
    SetControlEnaBled('CODEDUPLIQUER', False);
    SetControlProperty('CODEDUPLIQUER', 'TEXT', AncCode);
  end;
  if NomRubrique = 'JEU' then
  begin
    SetControlVisible('CODEDUPLIQUER', False);
    SetControlVisible('CODESPIN', True);
    CodeSpin := TSpinEdit(GetControl('CODESPIN'));
    if CodeSpin <> nil then CodeSpin.OnChange := AffectCode;
  end;

  Ecode := THEdit(GetControl('CODEDUPLIQUER'));
//If (ECode<>nil) and (NomRubrique<>'JEU') then Ecode.OnExit:=DefautCode;
  if (ECode <> nil) and (NomRubrique <> 'JEU') and (NomRubrique <> 'DUC') then
    Ecode.OnExit := DefautCode; // PT2

  MajFiltreNiveau(AncCode); //PT10

  EDate := THEdit(GetControl('DATEVALIDITE'));
  if Edate <> nil then Edate.OnDblClick := DateElipsisclick;

  Btn := TToolBarButton97(GetControl('BValider'));
  if Btn <> nil then Btn.OnClick := Validation;

  //DEB PT13
  THPredefini := THValComboBox(GetControl('PREDEFINI'));
  if THPredefini <> nil then THPredefini.OnChange := PredefiniChange;

  if GetControlText('PREDEFINI') = 'DOS' then
  begin
    SetControlText('CONVENTION','000');
    SetControlEnabled('CONVENTION',False);
  end
  else
  begin
    SetControlEnabled('CONVENTION',True);
    if ConvCol <> '' then
      SetControlText('CONVENTION',ConvCol); //PT15
  end;
  //FIN PT13
end;

procedure TOF_PGDUPLICATION.OnClose;
var
  ECode: THEdit;
begin
  inherited;
  Ecode := THEdit(GetControl('CODEDUPLIQUER'));
  if (ECode <> nil) then
  begin
    if Ecode.text = '' then
      DonneCodeDupliquer('', '', '','');//PT9
    if GetControlText ('PREDEFINI') <> PGCodePredefini then  // PT7
      if NomRubrique = 'PRO' then PGCodePredefini := GetControlText ('PREDEFINI'); // PT7
  end;
end;

procedure TOF_PGDUPLICATION.Validation(Sender: TObject);
var
  ECode, EDate: THEdit;
  EPred: THValComboBox;
  EConv: THValComboBox;
  Btn: TToolBarButton97;
  OKRub, OkElt, CEG, STD, DOS: Boolean;
  mes, CodeZero, temp, pred, cc, Conv: string;
begin
  CodeZero := StringOfChar('0', Longueur);
  Btn := TToolBarButton97(GetControl('BValider'));
  Ecode := THEdit(GetControl('CODEDUPLIQUER'));
  Epred := THValComboBox(GetControl('PREDEFINI'));
  EConv := THValComboBox(GetControl('CONVENTION')); //PT9
  EDate := THEdit(GetControl('DATEVALIDITE'));
  if Epred <> nil then pred := Epred.value;
  if EConv <> nil then Conv := EConv.value;
  if (Ecode <> nil) and (Epred <> nil) and (Btn <> nil) and (EDate <> nil) then
  begin
    DonneCodeDupliquer('', '', '','');//PT9
    if (ecode.text = '') and (pred <> '') and (EDate.text <> '') then
    begin
      PGIBox('Vous devez renseigner un nouveau code!', 'Duplication');
      exit;
    end;
    if (ecode.text <> '') and (pred <> '') and (EDate.text = '') then
    begin
      PGIBox('Vous devez renseigner une date de validité!', 'Duplication');
      exit;
    end;
    //DEB PT9
    if (ecode.text <> '') and (pred <> '') and (EConv.Visible = true) and (Conv = '') then
    begin
      PGIBox('Vous devez renseigner une convention!', 'Duplication');
      exit;
    end;
    //FIN PT9
    if (ecode.text <> '') and (pred <> '') and (EDate.text <> '') then
    begin
      AccesPredefini('TOUS', CEG, STD, DOS);
      if (Pred = 'CEG') and (CEG = FALSE) and (NomRubrique <> 'ELT') then
      begin
        PGIBox('Vous ne pouvez pas créer de rubrique prédéfinie CEGID', 'Accès refusé'); // PT5
        Pred := 'DOS';
        SetControlProperty('PREDEFINI', 'Value', 'DOS');
        exit;
      end;
      if (Pred = 'STD') and (STD = FALSE) and (NomRubrique <> 'ELT') then   
      begin
        PGIBox('Vous ne pouvez pas créer de rubrique prédéfinie Standard', 'Accès refusé'); // PT5
        Pred := 'DOS';
        SetControlProperty('PREDEFINI', 'Value', 'DOS');
        exit;
      end;
      //DEB PT11
      if (Pred = 'DOS') and (DOS = FALSE) then
      begin
        PGIBox('Vous ne pouvez pas créer de rubrique prédéfinie Dossier', 'Accès refusé');
        Pred := '';
        SetControlProperty('PREDEFINI', 'Value', 'DOS');
        exit;
      end;
      //FIN PT11

      if NomRubrique = 'ELT' then
      begin
// PT3 : 09/10/2002 : V585  PH  Duplication des elts nationaux autorisée avec le type DOS
        OkRub := TRUE;
//     OKRub:=TestRubriqueCegStd(pred,ecode.text,MaxRub);  Test supprimé
        OkElt := ExceptionEltNat(ecode.text);
        if OkElt then OkRub := OkElt;
      end
      else
// PT1 deb
        if (NomRubrique <> 'DUC') then
        begin
// PT1 fin
// DEB PT6
          if (NomRubrique = 'PRO') and (not IsNumeric(ecode.text)) then
          begin
            if (Length(Ecode.text) < longueur) then
            begin
              PGIBox('Vous devez saisir un code sur trois caractères.', Ecran.Caption);
              OkRub := FALSE;
            end
            else
            begin
              SetControlText('PREDEFINI', 'STD');
              SetControlEnabled('PREDEFINI', FALSE);
              OKRub := TRUE;
              if (not IsNumeric(ecode.text[3])) then
              begin
                PGIBox('Le dernier caractère doit être numérique.', Ecran.Caption);
                OkRub := FALSE;
              end;
              if (IsNumeric(ecode.text[1])) and (not IsNumeric(ecode.text[2])) then
              begin
                PGIBox('Vous devez saisir un code entièrement numérique ou composé de deux lettres suivies d''un chiffre.', Ecran.Caption);
                OkRub := FALSE;
              end;
            end;
          end
          else
            Begin
            if ecode.text[1] = 'P' then //PT12 gestion des variables de présence
              OKRub := TestRubrique(pred, RightStr(ecode.text,3), 200)
            else
              OKRub := TestRubrique(pred, ecode.text, MaxRub);
            { DEB PT8 }
            if ((NomRubrique = 'COT') OR (NomRubrique = 'AAA')) AND (not IsNumeric(ecode.text)) then
               Begin
                //PGIBox('Vous devez saisir un code numérique.', Ecran.Caption);
                OkRub := FALSE;
               End;
            { FIN PT8 }
            End;
        end
// FIN PT6
// PT1 deb
        else
          OKRub := TRUE;
// PT1 fin
      if (NomRubrique = 'VENT') then OkRub := True;
      if ((OkRub = False) or (ecode.text = CodeZero)) then //and (NomRubrique<>'ELT')then
      begin
        if NomRubrique = 'ELT' then Mes := MesTestRubrique('ELT', pred, MaxRub)
        else
// PT1 deb
          if (NomRubrique <> 'DUC') then
            Mes := MesTestRubrique('DUP', pred, MaxRub);
// PT1 fin
        LastError := 1;
        PGIBox(Mes, 'Code Erroné : ' + ecode.text);
        Ecode.text := '';
        if NomRubrique = 'JEU' then begin SetControlText('CODESPIN', '1'); Ecode.text := '1'; end;
        ECode.Enabled := True;
        Btn.enabled := True;
        SetFocusControl('CODEDUPLIQUER');
      end
      else
      begin
//PT1
        if (Length(Ecode.text) < longueur) and (NomRubrique <> 'DUC') then
//PT1 fin
        begin
          Temp := ColleZeroDevant(StrToInt(Ecode.text), Longueur);
          SetControlProperty('CODEDUPLIQUER', 'Text', Temp);
        end;
        PGIBox('Code correct.#13#10#13#10Cliquer sur la croix pour terminer le traitement.', 'Duplication');
        DonneCodeDupliquer(Ecode.text, Pred, EDate.text,Conv);   //PT9
        SetControlEnabled('CODESPIN', False);
        ECode.Enabled := False;
        Epred.Enabled := False;
        EDate.Enabled := False;
        Btn.enabled := False;
      end;
    end;
  end; //Begin nil
end;

//DEB PT10
procedure TOF_PGDUPLICATION.MajFiltreNiveau(CodeElt : String);
var
  Q : TQuery;
  {$IFNDEF EAGLCLIENT}
  Predefini : THDBValComboBox;
  {$ELSE}
  Predefini : THValComboBox;
  {$ENDIF}
  NiveauOK : String;
begin
  if CodeElt <> '' then
  begin
    NiveauOK := '"CEG","STD","DOS"';
    {$IFNDEF EAGLCLIENT}
    Predefini := THDBValComboBox(GetControl('PREDEFINI'));
    {$ELSE}
    Predefini := THValComboBox(GetControl('PREDEFINI'));
    {$ENDIF}
    if Predefini <> nil then
    begin
      Q := OpenSQL('SELECT PNR_TYPENIVEAU,PNR_NIVMAXPERSO FROM ELTNIVEAUREQUIS WHERE PNR_CODEELT="' + CodeElt + '" AND ##PNR_PREDEFINI## ORDER BY PNR_PREDEFINI DESC', True);
      if not Q.Eof then
      begin
        if (Q.Fields[1].AsString = 'CEG') then
          NiveauOK := '"CEG"';
        if (Q.Fields[1].AsString = 'STD') then
          NiveauOK := '"CEG","STD"';
      end;
      Ferme(Q);
    end;
    //PT11
    if (CEG = False) and (STD = False) then
      NiveauOK := '"DOS"';
    Predefini.Plus := ' AND CO_CODE IN (' + NiveauOK + ')';
  end;
end;
//FIN PT10

//DEB PT13
procedure TOF_PGDUPLICATION.PredefiniChange(Sender: TObject);
begin
  if GetControlText('PREDEFINI') = 'DOS' then
  begin
    SetControlText('CONVENTION','000');
    SetControlEnabled('CONVENTION',False);
  end
  else
    SetControlEnabled('CONVENTION',True);
end;
//FIN PT13

//DEB PT14
{ TOF_PGDUPLIETAB }
procedure TOF_PGDUPLIETAB.OnArgument(Arguments: string);
var
  AncCode,AncLib : String;
  Btn: TToolBarButton97;
begin
  DonneCodeDupliquerEtab('', '');
  AncCode := ReadTokenst(Arguments);
  AncLib := ReadTokenst(Arguments);

  SetControlProperty('CODEDUPLIQUER', 'TEXT', AncCode);
  SetControlProperty('LIBELLE', 'TEXT', AncLib);

  Btn := TToolBarButton97(GetControl('BValider'));
  if Btn <> nil then Btn.OnClick := Validation;
end;

procedure TOF_PGDUPLIETAB.OnClose;
var
  ECode,ELib: THEdit;
begin
  inherited;
  Ecode := THEdit(GetControl('CODEDUPLIQUER'));
  if (ECode <> nil) then
    if Ecode.text = '' then
      DonneCodeDupliquerEtab('', '');
  ELib := THEdit(GetControl('LIBELLE'));
  if (ELib <> nil) then
    if ELib.text = '' then
      DonneCodeDupliquerEtab('', '');
end;

procedure TOF_PGDUPLIETAB.Validation(Sender: TObject);
var
  ECode, ELib: THEdit;
  Btn: TToolBarButton97;
begin
  Btn := TToolBarButton97(GetControl('BValider'));
  Ecode := THEdit(GetControl('CODEDUPLIQUER'));
  ELib := THEdit(GetControl('LIBELLE'));
  if (Ecode <> nil) and (ELib <> nil) and (Btn <> nil) then
  begin
    DonneCodeDupliquer('', '', '','');
    if (ecode.text = '') then
    begin
      PGIBox('Vous devez renseigner un nouveau code établissement !', 'Duplication');
      SetFocusControl('CODEDUPLIQUER');
      exit;
    end;
    if (ELib.text = '') then
    begin
      PGIBox('Vous devez renseigner un nouveau libellé établissement !', 'Duplication');
      SetFocusControl('LIBELLE');
      exit;
    end;
    if (ecode.text <> '') and (ELib.text <> '') then
    begin
      if ExisteSQL('SELECT ET_ETABLISSEMENT FROM ETABLISS WHERE ET_ETABLISSEMENT="' + ecode.text + '"') then
      begin
        PGIBox('Le code établissement existe déjà');
        SetFocusControl('CODEDUPLIQUER');
        exit;
      end;
      if ExisteSQL('SELECT ET_ETABLISSEMENT FROM ETABLISS WHERE ET_LIBELLE="' + eLib.text + '"') then
      begin
        PGIBox('Le libellé établissement existe déjà');
        SetFocusControl('LIBELLE');
        exit;
      end;
      PGIBox('Code correct.#13#10#13#10Cliquer sur la croix pour terminer le traitement.', 'Duplication');
      DonneCodeDupliquerEtab(Ecode.text, ELib.text);
      ECode.Enabled := False;
      ELib.Enabled := False;
      Btn.enabled := False;
    end;
  end;
end;
//FIN PT14

//DEB PT16
{ TOF_PGDUPLIORG }
procedure TOF_PGDUPLIORG.OnArgument(Arguments: string);
var
  AncCode,AncLib : String;
  Btn: TToolBarButton97;
begin
  inherited;
  DonneCodeDupliquerEtab('', '');
  AncCode := ReadTokenst(Arguments);
  AncLib := ReadTokenst(Arguments);
  TypeDupli := READTOKENST(Arguments);
  EtabOri := AncLib;

  SetControlProperty('ORGINIT', 'TEXT', AncCode);
  SetControlEnabled('ORGINIT',False);

  if TypeDupli = '2' then
  begin
    SetControlProperty('METABLISSEMENT', 'TEXT', '<<Tous>>');
    SetControlVisible('ETABLISSEMENT',False);
    SetControlVisible('METABLISSEMENT',True);
    SetControlVisible('TNEWORG',False);
    SetControlVisible('NEWORG',False);
    SetControlProperty('NEWORG', 'VALUE', AncCode);
  end
  else
  begin
    SetControlProperty('ETABLISSEMENT', 'VALUE', AncLib);
    SetControlVisible('ETABLISSEMENT',True);
    SetControlVisible('METABLISSEMENT',False);
    SetControlVisible('TNEWORG',True);
    SetControlVisible('NEWORG',True);
  end;

  Btn := TToolBarButton97(GetControl('BValider'));
  if Btn <> nil then Btn.OnClick := Validation;
end;

procedure TOF_PGDUPLIORG.OnClose;
var
  ECode,ELib: THValComboBox;
  MEtab : THMultiValComboBox;
begin
  inherited;
  if TypeDupli <> '2' then
  begin
    Ecode := THValComboBox(GetControl('NEWORG'));
    if (ECode <> nil) then
      if Ecode.Value = '' then
        DonneCodeDupliquerEtab('', '');
    ELib := THValComboBox(GetControl('ETABLISSEMENT'));
    if (ELib <> nil) then
      if ELib.Value = '' then
        DonneCodeDupliquerEtab('', '');
  end;
end;

procedure TOF_PGDUPLIORG.Validation(Sender: TObject);
var
  ECode, ELib: THValComboBox;
  MEtab: THMultiValComboBox;
  Btn: TToolBarButton97;
  St, Etab, LEtab, LEtabErr: String;
  NbErr: integer;
  Tob_Org, TC, TOrg: TOB;
  Q: TQuery;
begin
  Btn := TToolBarButton97(GetControl('BValider'));
  Ecode := THValComboBox(GetControl('NEWORG'));
  if TypeDupli = '2' then
  begin
    MEtab := THMultiValComboBox(GetControl('METABLISSEMENT'));
    if (Ecode <> nil) and (MEtab <> nil) and (Btn <> nil) then
    begin
      LEtabErr := '';
      NbErr := 0;
      if (MEtab.text = '') then
      begin
        PGIBox('Vous devez renseigner au moins un établissement sur lequel dupliquer le code organisme !', 'Duplication');
        SetFocusControl('METABLISSEMENT');
        exit;
      end
      else
      begin
        St := 'SELECT * FROM ORGANISMEPAIE WHERE POG_ORGANISME="' + ecode.Value + '"' +
          ' AND POG_ETABLISSEMENT="' + EtabOri + '"';
        Tob_Org := Tob.Create('ORGANISMEPAIE',nil,-1);
        Tob_Org.LoadDetailDbFromSQL('ORGANISMEPAIE',st) ;
        TOrg := Tob_Org.Detail[0];
        if (MEtab.text = '<<Tous>>') then
        begin
          St := 'SELECT ET_ETABLISSEMENT FROM ETABLISS';
          Q := OpenSQL(St,True);
          LEtab := '';
          While not Q.Eof do
          begin
            if LEtab <> '' then
              LEtab := LEtab + ';' + Q.FindField('ET_ETABLISSEMENT').AsString
            else
              LEtab := Q.FindField('ET_ETABLISSEMENT').AsString;
            Q.Next;
          end;
          Ferme(Q);
        end
        else
          LEtab := MEtab.Text;

        while LEtab <> '' do
        begin
          Etab := READTOKENST(LEtab);
          if not ExisteSQL('SELECT POG_ORGANISME FROM ORGANISMEPAIE' +
            ' WHERE POG_ORGANISME="' + ECode.Value + '" AND POG_ETABLISSEMENT="' + Etab + '"') then
          begin
            TC := Tob.Create('ORGANISMEPAIE', Tob_Org, -1);
            TC.PutValue('POG_ETABLISSEMENT',Etab);
            TC.PutValue('POG_ORGANISME',TOrg.GetValue('POG_ORGANISME'));
            TC.PutValue('POG_LIBELLE',TOrg.GetValue('POG_LIBELLE'));
            TC.PutValue('POG_LIBEDITBULL',TOrg.GetValue('POG_LIBEDITBULL'));
            TC.PutValue('POG_NATUREORG',TOrg.GetValue('POG_NATUREORG'));
            TC.PutValue('POG_NUMAFFILIATION',TOrg.GetValue('POG_NUMAFFILIATION'));
            TC.PutValue('POG_PERIODICITDUCS',TOrg.GetValue('POG_PERIODICITDUCS'));
            TC.PutValue('POG_CAISSEDESTIN',TOrg.GetValue('POG_CAISSEDESTIN'));
            TC.PutValue('POG_AUTREPERIODUCS',TOrg.GetValue('POG_AUTREPERIODUCS'));
            TC.PutValue('POG_INSTITUTION',TOrg.GetValue('POG_INSTITUTION'));
            TC.PutValue('POG_NUMINTERNE',TOrg.GetValue('POG_NUMINTERNE'));
            TC.PutValue('POG_LGOPTIQUE',TOrg.GetValue('POG_LGOPTIQUE'));
            TC.PutValue('POG_DATECREATION',Date);
            TC.PutValue('POG_DATEMODIF',Date);
            TC.PutValue('POG_LONGTOTALE',TOrg.GetValue('POG_LONGTOTALE'));
            TC.PutValue('POG_LONGEDITABLE',TOrg.GetValue('POG_LONGEDITABLE'));
            TC.PutValue('POG_POSDEBUT',TOrg.GetValue('POG_POSDEBUT'));
            TC.PutValue('POG_BASETYPARR',TOrg.GetValue('POG_BASETYPARR'));
            TC.PutValue('POG_MTTYPARR',TOrg.GetValue('POG_MTTYPARR'));
            TC.PutValue('POG_TELEPHONE',TOrg.GetValue('POG_TELEPHONE'));
            TC.PutValue('POG_FAX',TOrg.GetValue('POG_FAX'));
            TC.PutValue('POG_TELEX',TOrg.GetValue('POG_TELEX'));
            TC.PutValue('POG_SIRET',TOrg.GetValue('POG_SIRET'));
            TC.PutValue('POG_CONTACT',TOrg.GetValue('POG_CONTACT'));
            TC.PutValue('POG_REGROUPEMENT',TOrg.GetValue('POG_REGROUPEMENT'));
            TC.PutValue('POG_RUPTSIRET',TOrg.GetValue('POG_RUPTSIRET'));
            TC.PutValue('POG_RUPTAPE',TOrg.GetValue('POG_RUPTAPE'));
            TC.PutValue('POG_RUPTGROUPE',TOrg.GetValue('POG_RUPTGROUPE'));
            TC.PutValue('POG_REGROUPDADSU',TOrg.GetValue('POG_REGROUPDADSU'));
            TC.PutValue('POG_PREVOYANCE',TOrg.GetValue('POG_PREVOYANCE'));
            TC.PutValue('POG_EMAIL',TOrg.GetValue('POG_EMAIL'));
            TC.PutValue('POG_RUPTNUMERO',TOrg.GetValue('POG_RUPTNUMERO'));
            TC.PutValue('POG_ADRESSE1',TOrg.GetValue('POG_ADRESSE1'));
            TC.PutValue('POG_ADRESSE2',TOrg.GetValue('POG_ADRESSE2'));
            TC.PutValue('POG_ADRESSE3',TOrg.GetValue('POG_ADRESSE3'));
            TC.PutValue('POG_CODEPOSTAL',TOrg.GetValue('POG_CODEPOSTAL'));
            TC.PutValue('POG_VILLE',TOrg.GetValue('POG_VILLE'));
            TC.PutValue('POG_GENERAL',TOrg.GetValue('POG_GENERAL'));
            TC.PutValue('POG_CODAPPLI',TOrg.GetValue('POG_CODAPPLI'));
            TC.PutValue('POG_SERVUNIQ',TOrg.GetValue('POG_SERVUNIQ'));
            TC.PutValue('POG_SOUSTOTDUCS',TOrg.GetValue('POG_SOUSTOTDUCS'));
            TC.PutValue('POG_POSTOTAL',TOrg.GetValue('POG_POSTOTAL'));
            TC.PutValue('POG_LONGTOTAL',TOrg.GetValue('POG_LONGTOTAL'));
            TC.PutValue('POG_PAIEGROUPE',TOrg.GetValue('POG_PAIEGROUPE'));
            TC.PutValue('POG_RIBDUCSEDI',TOrg.GetValue('POG_RIBDUCSEDI'));
            TC.PutValue('POG_PAIEMODE',TOrg.GetValue('POG_PAIEMODE'));
            TC.PutValue('POG_IDENTOPS',TOrg.GetValue('POG_IDENTOPS'));
            TC.PutValue('POG_NOCONTEMET',TOrg.GetValue('POG_NOCONTEMET'));
            TC.PutValue('POG_CENTREPAYEUR',TOrg.GetValue('POG_CENTREPAYEUR'));
            TC.PutValue('POG_CPINFO',TOrg.GetValue('POG_CPINFO'));
            TC.PutValue('POG_DUCSDOSSIER',TOrg.GetValue('POG_DUCSDOSSIER'));
            TC.PutValue('POG_PERIODCALCUL',TOrg.GetValue('POG_PERIODCALCUL'));
            TC.PutValue('POG_AUTPERCALCUL',TOrg.GetValue('POG_AUTPERCALCUL'));
            TC.PutValue('POG_DUCSREGLEMENT',TOrg.GetValue('POG_DUCSREGLEMENT'));
            TC.PutValue('POG_DUCEXIGIBILITE',TOrg.GetValue('POG_DUCEXIGIBILITE'));
            TC.PutValue('POG_DUCLIMITEDEPOT',TOrg.GetValue('POG_DUCLIMITEDEPOT'));
            TC.PutValue('POG_VCPA',TOrg.GetValue('POG_VCPA'));
            TC.PutValue('POG_IDENTQUAL',TOrg.GetValue('POG_IDENTQUAL'));
            TC.PutValue('POG_IDENTEMET',TOrg.GetValue('POG_IDENTEMET'));
            TC.PutValue('POG_IDENTDEST',TOrg.GetValue('POG_IDENTDEST'));
            TC.PutValue('POG_ADHERCONTACT',TOrg.GetValue('POG_ADHERCONTACT'));
            TC.PutValue('POG_BLOCNOTE',TOrg.GetValue('POG_BLOCNOTE'));
            TC.PutValue('POG_TYPEPERIOD',TOrg.GetValue('POG_TYPEPERIOD'));
            TC.PutValue('POG_TYPEAUTPERIOD',TOrg.GetValue('POG_TYPEAUTPERIOD'));
            TC.PutValue('POG_TITULAIRECPT',TOrg.GetValue('POG_TITULAIRECPT'));
            TC.PutValue('POG_CONDSPEC',TOrg.GetValue('POG_CONDSPEC'));
            TC.PutValue('POG_PGORGGUID',TOrg.GetValue('POG_PGORGGUID'));
            TC.InsertOrUpdateDB();
          end
          else
          begin
            NbErr := NbErr + 1;
            if LEtabErr <> '' then
              LEtabErr := LEtabErr + ';' + Etab
            else
              LEtabErr := Etab;
          end;
        end;
      end;
      if LEtabErr <> '' then
      begin
        if NbErr = 1 then
          PGIBox('La duplication a été effectuée.#13#10#13#10L''établissement '+ LEtabErr + ' existait déjà.#13#10#13#10Cliquer sur la croix pour terminer le traitement.', 'Duplication')
        else
          PGIBox('La duplication a été effectuée.#13#10#13#10Les établissements '+ LEtabErr + ' existaient déjà.#13#10#13#10Cliquer sur la croix pour terminer le traitement.', 'Duplication');
      end
      else
        PGIBox('La duplication a été effectuée.#13#10#13#10Cliquer sur la croix pour terminer le traitement.', 'Duplication');
      ECode.Enabled := False;
      MEtab.Enabled := False;
      Btn.enabled := False;
    end;
  end
  else
  begin
    ELib := THValComboBox(GetControl('ETABLISSEMENT'));
    if (Ecode <> nil) and (ELib <> nil) and (Btn <> nil) then
    begin
      DonneCodeDupliquer('', '', '','');
      if (ecode.text = '') then
      begin
        PGIBox('Vous devez renseigner un nouveau code organisme !', 'Duplication');
        SetFocusControl('NEWORG');
        exit;
      end;
      if (ELib.text = '') then
      begin
        PGIBox('Vous devez renseigner l''établissement !', 'Duplication');
        SetFocusControl('ETABLISSEMENT');
        exit;
      end;
      if (ecode.Value <> '') and (ELib.Value <> '') then
      begin
        if ExisteSQL('SELECT POG_ORGANISME FROM ORGANISMEPAIE' +
          ' WHERE POG_ORGANISME="' + ecode.Value + '" AND POG_ETABLISSEMENT="' + ELib.Value + '"') then
        begin
          PGIBox('Ce code organisme existe déjà pour cet établissement');
          SetFocusControl('NEWORG');
          exit;
        end;
      end;
      PGIBox('Code correct.#13#10#13#10Cliquer sur la croix pour terminer le traitement.', 'Duplication');
      DonneCodeDupliquerEtab(Ecode.Value, ELib.Value);
      ECode.Enabled := False;
      ELib.Enabled := False;
      Btn.enabled := False;
    end;
  end;
end;
//FIN PT16

initialization
  registerclasses([TOF_PGDUPLICATION,TOF_PGDUPLIETAB,TOF_PGDUPLIORG]);
end.

