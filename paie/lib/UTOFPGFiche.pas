{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 13/09/2001
Modifié le ... :   /  /
Description .. : Gestion de plusieurs types de  fiches
Mots clefs ... : PAIE;FICHE
*****************************************************************
PT1 18/04/2002 SB V571 Fiche de bug n°10087 : V_PGI_env.LibDossier non renseigné en Mono
PT2 21/01/2003 PH V591 Fiche de bug n°10443 : Ergonomie Entete ecran
PT3 09/02/2007 FC V_70 Rajout de la classe PGMULELENATDOS pour la gestion des éléments nationaux Dossier
PT4 13/03/2007 FC V_70 Rajout de la classe PGMULELTNIVREQUIS pour la gestion des niveaux préconisés des éléments nationaux
PT5 06/06/2007 FC V_72 Passer en argument les critères du MUL à la synthèse des éléments
PT6 14/06/2007 FC V_72 Ne pas proposer le niveau Population si elles ne sont pas gérées
PT7 04/07/2007 FC V_72 FQ 14475 Mise en place des concepts
PT8 13/07/2007 FC V_72 Interdire le double click si non accès au btn insert
PT9 03/09/2007 FC V_80 FQ 14722 prise en compte du prédéfini CEG
}
unit UTOFPGFiche;

interface

uses Controls, Classes, HCtrls,
  ParamSoc,
{$IFDEF EAGLCLIENT}
  eMul,MaineAgl,Utob,
{$ELSE}
  Mul,Fe_Main,HDB,
{$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  UTOF,P5Def,HTB97,Entpaie,HEnt1,sysutils,StrUtils
{$IFNDEF CPS1}
  ,PGPOPULOUTILS
{$ENDIF}
  , PgOutils,HQry,ComCtrls,HMsgBox  //PT11
 ;

type
  TOF_PGMULCOEFF = class(TOF)
    procedure OnArgument(Arguments: string); override;
  end;
type
  TOF_PGMULAGE = class(TOF)
    procedure OnArgument(Arguments: string); override;
  end;
type
  TOF_PGMULCOTISATION = class(TOF)
    procedure OnArgument(Arguments: string); override;
  end;
type
  TOF_PGMULELENAT = class(TOF)
    procedure OnArgument(Arguments: string); override;
  private
    procedure AffichageSynthese(Sender : TObject); //PT5
  end;
type
  TOF_PGMULMINICONVENT = class(TOF)
    procedure OnArgument(Arguments: string); override;
  end;
type
  TOF_PGMULREMUNERATION = class(TOF)
    procedure OnArgument(Arguments: string); override;
  end;
type
  TOF_PGVARIABLEETAT = class(TOF)
    procedure OnArgument(Arguments: string); override;
  end;
type
  TOF_PGPROFILETAT = class(TOF)
    procedure OnArgument(Arguments: string); override;
  end;
//DEB PT3
type
  TOF_PGMULELENATDOS = class(TOF)
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;
  private
    ParamTyp:String;
    ParamVal:String;
    Parametres : String;///PT8
    NonAcces : Boolean;//PT8
    procedure ChangeXXWhere2;
    procedure ChangeValeurType(Sender : TObject);
    procedure ChangeValeurSaisie(Sender : TObject);
    procedure ExitValeurSaisie(Sender : TObject);
    procedure AjoutEnregistrement(Sender : TObject);
    procedure GrilleDblClick(Sender: TObject);
  end;
//FIN PT3
//DEB PT4
type
  TOF_PGMULELTNIVREQUIS = class(TOF)
  public
    procedure OnArgument(Arguments: string); override;
  end;
//FIN PT4

implementation
uses PgOutils2,  PGEditOutils2, DBGrids, DB;
{ TOF_PGMULCOEFF }

procedure TOF_PGMULCOEFF.OnArgument(Arguments: string);
var
  Combo: TControl;
begin
  Combo := ThValComboBox(getcontrol('PMI_NATURE'));
  InitialiseCombo(Combo);
  Combo := ThValComboBox(getcontrol('PMI_CONVENTION'));
  InitialiseCombo(Combo);
end;

{ TOF_PGMULAGE }

procedure TOF_PGMULAGE.OnArgument(Arguments: string);
var
  Combo: TControl;
begin
  Combo := ThValComboBox(getcontrol('PMI_CONVENTION'));
  InitialiseCombo(Combo);
end;

{ TOF_PGMULCOTISATION }

procedure TOF_PGMULCOTISATION.OnArgument(Arguments: string);
var
  Combo: TControl;
begin
  Combo := ThValComboBox(getcontrol('PCT_THEMECOT'));
  InitialiseCombo(Combo);
  Combo := ThValComboBox(getcontrol('PCT_NATURERUB'));
  InitialiseCombo(Combo);
end;

{ TOF_PGMULELENAT }

procedure TOF_PGMULELENAT.OnArgument(Arguments: string);
var
  Combo: TControl;
  Btn : TToolbarButton97;
begin
  Combo := ThValComboBox(getcontrol('PEL_THEMEELT'));
  InitialiseCombo(Combo);

  //DEB PT3
  if not GetParamSocSecur('SO_PGGESTELTDYNDOS',False) then
  begin
    SetControlVisible('BSYNTHESE',False);
    SetControlVisible('BNIVEAUREQUIS',False);
  end;
  //FIN PT3

  Btn := TToolbarButton97(getcontrol('BSYNTHESE'));
  if btn <> nil then Btn.OnClick := AffichageSynthese;

  if not JaiLeDroitTag(41104) then
    SetControlEnabled('BNIVEAUREQUIS',False);

  PaieConceptPlanPaie(Ecran);   //PT7
end;

//DEB PT5
procedure TOF_PGMULELENAT.AffichageSynthese(Sender : TObject);
begin
  AglLanceFiche('PAY', 'SYNELTNAT_MUL', '', '', GetControlText('PEL_CODEELT') + ';' + GetControlText('PEL_THEMEELT') + ';' + GetControlText('PEL_DATEVALIDITE') + ';' + GetControlText('PEL_DATEVALIDITE_'));
end;
//FIN PT5

{ TOF_PGMULMINICONVENT }

procedure TOF_PGMULMINICONVENT.OnArgument(Arguments: string);
var
  Combo: TControl;
begin
  Combo := ThValComboBox(getcontrol('PMI_CONVENTION'));
  InitialiseCombo(Combo);
  //PT2 21/01/2003 PH V591 Fiche de bug n°10443 : Ergonomie Entete ecran
  if Arguments = 'C' then Ecran.caption := 'Coefficient'
  else if Arguments = 'I' then Ecran.caption := 'Indice'
  else if Arguments = 'N' then Ecran.caption := 'Niveau'
  else if Arguments = 'Q' then Ecran.caption := 'Qualification';
  UpdateCaption(TFMul(Ecran));
  // FIN PT2
end;

{ TOF_PGMULREMUNERATION }

procedure TOF_PGMULREMUNERATION.OnArgument(Arguments: string);
var
  Combo: TControl;
begin
  Combo := ThValComboBox(getcontrol('PRM_THEMEREM'));
  InitialiseCombo(Combo);
end;

{ TOF_PGVARIABLEETAT }

procedure TOF_PGVARIABLEETAT.OnArgument(Arguments: string);
var
  Combo: TControl;
  Defaut: ThEdit;
  Min, Max: string;
begin
  //Affectation des Valeurs par défaut
  Defaut := ThEdit(getcontrol('DOSSIER'));
  if Defaut <> nil then
    //  Defaut.text:=V_PGI_env.LibDossier;  //PT1 Mise en commentaire
    Defaut.text := GetParamSoc('SO_LIBELLE');
  RecupMinMaxTablette('PG', 'VARIABLEPAIE', 'PVA_VARIABLE', Min, Max);
  Defaut := ThEdit(getcontrol('PVA_VARIABLE'));
  if Defaut <> nil then Defaut.text := Min;
  Defaut := ThEdit(getcontrol('PVA_VARIABLE_'));
  if Defaut <> nil then Defaut.text := Max;
  Combo := ThValComboBox(getcontrol('PVA_NATUREVAR'));
  InitialiseCombo(Combo);
  Combo := ThValComboBox(getcontrol('PVA_THEMEVAR'));
  InitialiseCombo(Combo);
  Combo := ThValComboBox(getcontrol('PVA_PREDEFINI'));
  InitialiseCombo(Combo);

end;

{ TOF_PGPROFILETAT }


procedure TOF_PGPROFILETAT.OnArgument(Arguments: string);
var
  Combo: TControl;
  Defaut: ThEdit;
  Min, Max: string;
begin
  inherited;
  //Affectation des Valeurs par défaut
  Defaut := ThEdit(getcontrol('DOSSIER'));
  if Defaut <> nil then
    //  Defaut.text:=V_PGI_env.LibDossier;// PT1 Mise en commentaire
    Defaut.text := GetParamSoc('SO_LIBELLE');
  RecupMinMaxTablette('PG', 'PROFILPAIE', 'PPI_PROFIL', Min, Max);
  Defaut := ThEdit(getcontrol('PPI_PROFIL'));
  if Defaut <> nil then Defaut.text := Min;
  Defaut := ThEdit(getcontrol('PPI_PROFIL_'));
  if Defaut <> nil then Defaut.text := Max;
  Combo := ThValComboBox(getcontrol('PPI_TYPEPROFIL'));
  InitialiseCombo(Combo);
  Combo := ThValComboBox(getcontrol('PPI_THEMEPROFIL'));
  InitialiseCombo(Combo);
  Combo := ThValComboBox(getcontrol('PPM_NATURERUB'));
  InitialiseCombo(Combo);
  Combo := ThValComboBox(getcontrol('PPI_PREDEFINI'));
  InitialiseCombo(Combo);
end;

//DEB PT3
{ TOF_PGMULELENATDOS }
procedure TOF_PGMULELENATDOS.OnArgument(Arguments: string);
var
  Combo: TControl;
  CmbType: ThValComboBox;
  Arg:String;
  Btn: TToolBarButton97;
  Param:THEdit;
  Etab, Sal, Pop : THEdit;
  LibValeur:THLabel;
  LibelleValeur:THLabel;
  Q : TQuery;
  CodeElt:THEdit;
  Predefini:String;
  CEG, STD, DOS : Boolean;//PT7
begin
  Arg := Arguments;
  Parametres := Arguments;
  ParamTyp := '';
  ParamVal := '';
  if Arg <> '' then
  begin
    Param := THEdit(getcontrol('PARAM'));
    if Param <> nil then
      Param.Text := Arg;
    ParamTyp := READTOKENST(Arg);
    ParamVal := READTOKENST(Arg);
  end;

  //DEB PT6
{$IFNDEF CPS1}
  //S'il existe un parametre population
  if not ExisteSQL('SELECT PPC_POPULATION FROM ORDREPOPULATION') then
  begin
    CmbType := ThValComboBox(getcontrol('PED_TYPENIVEAU'));
    if CmbType <> nil then
      CmbType.Plus := ' AND CO_CODE <> "POP"';
  end;
{$ENDIF}
  //FIN PT6

  Btn := TToolBarButton97(getcontrol('BInsert'));
  if Btn <> nil then
    Btn.Onclick := AjoutEnregistrement;

  TFMul(Ecran).FListe.OnDblClick := GrilleDblClick; //PT8

  CmbType := ThValComboBox(getcontrol('PED_TYPENIVEAU'));
  if (CmbType <> nil) then
    CmbType.OnChange := ChangeValeurType;

  Etab := THEdit(getcontrol('PED_ETABLISSEMENT'));
  if (Etab<>nil) then
  begin
    Etab.OnChange := ChangeValeurSaisie;
    Etab.OnExit := ExitValeurSaisie;
  end;

  Sal := THEdit(getcontrol('PED_SALARIE'));
  if (Sal<>nil) then
  begin
    Sal.OnChange := ChangeValeurSaisie;
    Sal.OnExit := ExitValeurSaisie;
  end;

  Pop := THEdit(getcontrol('PED_CODEPOP'));
  if (Pop<>nil) then
  begin
    Pop.OnChange := ChangeValeurSaisie;
    Pop.OnExit := ExitValeurSaisie;
{$IFNDEF CPS1}
    Predefini := GetPredefiniPopulation('PAI');
    Pop.Plus := ' AND PPC_PREDEFINI = "' + Predefini + '" ';
{$ENDIF}
  end;

  if (ParamTyp = 'ELT') then
  begin
    CodeElt := THEdit(getcontrol('PED_CODEELT'));
    if CodeElt <> nil then
    begin
      CodeElt.Text := ParamVal;
      if ParamVal <> '' then
        CodeElt.Enabled := False;
    end;
    SetControlVisible('PED_ETABLISSEMENT',False);
    SetControlVisible('PED_SALARIE',False);
    SetControlVisible('PED_CODEPOP',False);
    SetControlVisible('TVALEUR',False);
    TFMul(Ecran).SetDBListe('PGELTNATDOS');
  end
  else
  begin
    Combo := ThValComboBox(getcontrol('PED_THEMEELT'));
    InitialiseCombo(Combo);
    Combo := ThValComboBox(getcontrol('PED_TYPENIVEAU'));
    InitialiseCombo(Combo);

    LibValeur := THLabel(getcontrol('TVALEUR'));
    if (ParamTyp <> '') then
    begin
      SetControlVisible('TVALEUR',True);
      SetControlText('PED_TYPENIVEAU',ParamTyp);
      if (ParamTyp = 'ETB') then
      begin
        SetControlVisible('PED_ETABLISSEMENT',True);
        SetControlVisible('PED_SALARIE',False);
        SetControlVisible('PED_CODEPOP',False);
        SetControlText('PED_ETABLISSEMENT',ParamVal);
        SetControlEnabled('PED_ETABLISSEMENT',False);
        LibValeur.Caption := 'Etablissement';
        Q := OpenSQL('SELECT ET_LIBELLE FROM ETABLISS WHERE ET_ETABLISSEMENT="' + ParamVal + '"', True);
      end;
      if (ParamTyp = 'SAL') then
      begin
        SetControlVisible('PED_ETABLISSEMENT',False);
        SetControlVisible('PED_SALARIE',True);
        SetControlVisible('PED_CODEPOP',False);
        SetControlText('PED_SALARIE',ParamVal);
        SetControlEnabled('PED_SALARIE',False);
        LibValeur.Caption := 'Salarié';
        Q := OpenSQL('SELECT PSA_LIBELLE||" "||PSA_PRENOM FROM SALARIES WHERE PSA_SALARIE="' + ParamVal + '"', True);
      end;
      if (ParamTyp = 'POP') then
      begin
        SetControlVisible('PED_ETABLISSEMENT',False);
        SetControlVisible('PED_SALARIE',False);
        SetControlVisible('PED_CODEPOP',True);
        SetControlText('PED_CODEPOP',ParamVal);
        SetControlEnabled('PED_CODEPOP',False);
        LibValeur.Caption := 'Population';
        Q := OpenSQL('SELECT PPC_LIBELLE FROM ORDREPOPULATION WHERE PPC_POPULATION="' + ParamVal + '"', True);
      end;

      // Afficher le libellé de la valeur du niveau
      LibelleValeur := THLabel(GetControl('LBL_VALEURNIVEAU'));
      if (LibelleValeur <> nil) then
        if not Q.Eof then
          LibelleValeur.Caption := Q.Fields[0].AsString
        else
          LibelleValeur.Caption := '';
      ferme(Q);

      CmbType.Enabled := False;
    end
    else
    begin
      SetControlVisible('PED_ETABLISSEMENT',False);
      SetControlVisible('PED_SALARIE',False);
      SetControlVisible('PED_CODEPOP',False);
      SetControlVisible('TVALEUR',False);
      CmbType.Enabled := True;
      LibValeur.Caption := '';
    end;
  end;

//  ChangeXXWhere2;

  //DEB PT7
  AccesPredefini('TOUS', CEG, STD, DOS);
{$IFDEF CPS1}
   STD := FALSE;
{$ENDIF}
  NonAcces := False; //PT8
  if (DOS = FALSE) then // or (not ExisteSQL('SELECT PNR_CODEELT FROM ELTNIVEAUREQUIS WHERE PNR_CODEELT = "' + GetControlText('PED_CODEELT') + '" AND PNR_NIVMAXPERSO IN ("ETB","POP","SAL") AND ##PNR_PREDEFINI##')) then
  begin
    SetControlEnabled('BInsert',False);
    NonAcces := True;  //PT8
  end;
  //FIN PT7
end;

procedure TOF_PGMULELENATDOS.ChangeValeurType(Sender : TObject);
var
  TypNiveau:THValComboBox;
  Liste : String;
begin
  TypNiveau := THValComboBox(GetControl('PED_TYPENIVEAU'));

  if TypNiveau.value = '' then
  begin
    SetControlVisible('PED_ETABLISSEMENT',False);
    SetControlVisible('PED_SALARIE',False);
    SetControlVisible('PED_CODEPOP',False);
    SetControlVisible('TVALEUR',False);
    SetControlText('TVALEUR','');
    SetControlText('LBL_VALEURNIVEAU','');
    SetControlText('PED_ETABLISSEMENT','');
    SetControlText('PED_SALARIE','');
    SetControlText('PED_CODEPOP','');
    Liste := 'PGELTNATDOS';
  end;

  if TypNiveau.value = 'ETB' then
  begin
    SetControlVisible('PED_ETABLISSEMENT',True);
    SetControlVisible('PED_SALARIE',False);
    SetControlVisible('PED_CODEPOP',False);
    SetControlVisible('TVALEUR',True);
    SetControlText('TVALEUR','Etablissement');
    SetControlText('LBL_VALEURNIVEAU','');
    SetControlText('PED_ETABLISSEMENT','');
    SetControlText('PED_SALARIE','');
    SetControlText('PED_CODEPOP','');
    Liste := 'PGELTNATDOSETB';
  end;

  if TypNiveau.value = 'SAL' then
  begin
    SetControlVisible('PED_ETABLISSEMENT',False);
    SetControlVisible('PED_SALARIE',True);
    SetControlVisible('PED_CODEPOP',False);
    SetControlVisible('TVALEUR',True);
    SetControlText('TVALEUR','Salarié');
    SetControlText('LBL_VALEURNIVEAU','');
    SetControlText('PED_ETABLISSEMENT','');
    SetControlText('PED_SALARIE','');
    SetControlText('PED_CODEPOP','');
    Liste := 'PGELTNATDOSSAL';
  end;

  if TypNiveau.value = 'POP' then
  begin
    SetControlVisible('PED_ETABLISSEMENT',False);
    SetControlVisible('PED_SALARIE',False);
    SetControlVisible('PED_CODEPOP',True);
    SetControlVisible('TVALEUR',True);
    SetControlText('TVALEUR','Population');
    SetControlText('LBL_VALEURNIVEAU','');
    SetControlText('PED_ETABLISSEMENT','');
    SetControlText('PED_SALARIE','');
    SetControlText('PED_CODEPOP','');
    Liste := 'PGELTNATDOSPOP';
  end;

  TFMul(Ecran).SetDBListe(Liste);
  TFMul(Ecran).BCherche.Click;
end;

procedure TOF_PGMULELENATDOS.OnLoad;
begin
  inherited;
  ChangeXXWhere2;
end;

procedure TOF_PGMULELENATDOS.ChangeXXWhere2;
var
  Where2:THEdit;
  Habilitation:String;
  St:String;
  Etab:String;
  j:integer;
begin
  Habilitation := '';
  Where2 := THEdit(GetControl('XX_WHERE2'));

  if Assigned(MonHabilitation) then
  begin
    if (MonHabilitation.LesEtab <> '') then
    begin
      St := MonHabilitation.LesEtab;
      Etab := ReadTokenSt(St);
      j := 0;
      while Etab <> '' do
      begin
        j := j + 1;
        if Etab <> '' then
        begin
          if j > 1 then Habilitation := Habilitation + ',';
          Habilitation := Habilitation + '"' + Etab + '"';
        end;
        Etab := ReadTokenSt(St);
      end;
      if j > 0 then
        Habilitation := '(PED_ETABLISSEMENT IN (' + Habilitation + '))';
    end
    else
      Habilitation := '(PED_ETABLISSEMENT <> "")';

    if (MonHabilitation.LeSQL <> '') then
    begin
      if Habilitation <> '' then
        Habilitation := Habilitation + ' OR (' + MonHabilitation.LeSQL + ')'
      else
        Habilitation := '(' + MonHabilitation.LeSQL + ')';
    end
    else
      if Habilitation <> '' then
        Habilitation := Habilitation + ' OR (PED_SALARIE <> "")'
      else
        Habilitation := '(PED_SALARIE <> "")';

    if Habilitation <> '' then
      Habilitation := Habilitation + ' OR (PED_CODEPOP <> "")'
    else
      Habilitation := '(PED_CODEPOP <> "")';
  end;

  if (Where2 <> nil) then SetControlText('XX_WHERE2', Habilitation);
end;

procedure TOF_PGMULELENATDOS.AjoutEnregistrement(Sender : TObject);
begin
  if (ParamTyp <> '') and (ParamVal <> '') then
    if ParamTyp <> 'ELT' then
      AglLanceFiche('PAY', 'ELEMENTNATDOS', '', '', 'ACTION=CREATION;' + ParamTyp + ';' + ParamVal)
    else
      AglLanceFiche('PAY', 'ELEMENTNATDOS', '', '', 'ACTION=CREATION;' + ';;' + ParamVal)
  else
    AglLanceFiche('PAY', 'ELEMENTNATDOS', '', '', 'ACTION=CREATION');
  TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

procedure TOF_PGMULELENATDOS.ChangeValeurSaisie(Sender: TObject);
var
  LibValeur:THLabel;
begin
  LibValeur := THLabel(getcontrol('LBL_VALEURNIVEAU'));
  if (LibValeur<>nil) then
    LibValeur.Caption := '';
end;

procedure TOF_PGMULELENATDOS.ExitValeurSaisie(Sender: TObject);
var
//  ValeurSaisie:THEdit;
  LibValeur:THLabel;
  Q:TQuery;
  Salarie : String;
  TypNiveau : THValComboBox;
begin
  TypNiveau := THValComboBox(GetControl('PED_TYPENIVEAU'));
  // Afficher le libellé de la valeur qui vient d'être récupérée
  if (TypNiveau.value = 'ETB') then
    Q := OpenSQL('SELECT ET_LIBELLE FROM ETABLISS WHERE ET_ETABLISSEMENT="' + GetControlText('PED_ETABLISSEMENT') + '"', True);
  if (TypNiveau.value = 'SAL') then
  begin
    Salarie := GetControlText('PED_SALARIE');
    if (VH_PAIE.PgTypeNumSal = 'NUM') and (isnumeric(Salarie)) and (Salarie <> '') then
    begin
       Salarie := ColleZeroDevant(StrToInt(trim(Salarie)), 10);
       SetControlText('PED_SALARIE', Salarie);
    end;
    Q := OpenSQL('SELECT PSA_LIBELLE||" "||PSA_PRENOM FROM SALARIES WHERE PSA_SALARIE="' + Salarie + '"', True);
  end;
  if (TypNiveau.value = 'POP') then
    Q := OpenSQL('SELECT PPC_LIBELLE FROM ORDREPOPULATION WHERE PPC_POPULATION="' + GetControlText('PED_CODEPOP') + '"', True);

  LibValeur := THLabel(getcontrol('LBL_VALEURNIVEAU'));
  if (LibValeur<>nil) then
    if not Q.Eof then
      LibValeur.Caption := Q.Fields[0].AsString
    else
      LibValeur.Caption := '';
  Ferme(Q);
end;
//FIN PT3
//DEB PT8
procedure TOF_PGMULELENATDOS.GrilleDblClick(Sender: TObject);
var
  PTypNiv, PValNiv, PCodElt, PDatVal :String;
begin
  if not NonAcces then
  begin
    if (TFMUL(Ecran).Q.RecordCount <> 0) then
    begin
      PTypNiv := TFMUL(Ecran).Q.Findfield('PED_TYPENIVEAU').asstring;
      PValNiv := TFMUL(Ecran).Q.Findfield('PED_VALEURNIVEAU').asstring;
      PCodElt := TFMUL(Ecran).Q.Findfield('PED_CODEELT').asstring;
      PDatVal := TFMUL(Ecran).Q.Findfield('PED_DATEVALIDITE').AsVariant;
      AglLanceFiche('PAY','ELEMENTNATDOS','',PTypNiv+';'+PValNiv+';'+PCodElt+';'+PDatVal,'ACTION=MODIF;'+Parametres)
    end
    else
    begin
      if (ParamTyp <> '') and (ParamVal <> '') then
        if ParamTyp <> 'ELT' then
          AglLanceFiche('PAY', 'ELEMENTNATDOS', '', '', 'ACTION=CREATION;' + ParamTyp + ';' + ParamVal)
        else
          AglLanceFiche('PAY', 'ELEMENTNATDOS', '', '', 'ACTION=CREATION;' + ';;' + ParamVal)
      else
        AglLanceFiche('PAY', 'ELEMENTNATDOS', '', '', 'ACTION=CREATION');
    end;
    TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
  end;
end;
//FIN PT8

//DEB PT4
{ TOF_PGMULELTNIVREQUIS }
procedure TOF_PGMULELTNIVREQUIS.OnArgument(Arguments: string);
var
  Niveau : ThValComboBox;
  Plus,StPop:String;
  CEG, STD, DOS : Boolean;//PT7
begin
  //DEB PT6
{$IFNDEF CPS1}
  //S'il existe un parametre population
  if not ExisteSQL('SELECT PPC_POPULATION FROM ORDREPOPULATION') then
    StPop := '"POP"'
  else
    StPop := '';
{$ENDIF}
  //FIN PT6

  if not GetParamSocSecur('SO_PGGESTELTDYNDOS',False) then
  begin
    Plus := ' AND CO_CODE <> "ETB" AND CO_CODE <> "POP" AND CO_CODE <> "SAL"';
    Niveau := ThValComboBox(getcontrol('PNR_TYPENIVEAU'));
    if Niveau <> nil then
      Niveau.Plus := Plus;
  end
  //DEB PT6
  else if StPop <> '' then
  begin
    Plus := ' AND CO_CODE <> "POP"';
    Niveau := ThValComboBox(getcontrol('PNR_TYPENIVEAU'));
    if Niveau <> nil then
      Niveau.Plus := Plus;
  end;
  //FIN PT6

  //DEB PT9
  Plus := ' AND CO_CODE <> "DOS"';
  Niveau := ThValComboBox(getcontrol('PNR_PREDEFINI'));
  if Niveau <> nil then
    Niveau.Plus := Plus;
  //FIN PT9

  PaieConceptPlanPaie(Ecran); //PT7

//  SetControlText('PNR_PREDEFINI','STD');    //PT9
//  SetControlEnabled('PNR_PREDEFINI',False);    //PT9
end;
//FIN PT4

initialization
  registerclasses([TOF_PGMULCOEFF, TOF_PGMULAGE, TOF_PGMULCOTISATION, TOF_PGMULELENAT, TOF_PGMULMINICONVENT, TOF_PGMULREMUNERATION, TOF_PGVARIABLEETAT, TOF_PGPROFILETAT, TOF_PGMULELENATDOS, TOF_PGMULELTNIVREQUIS]);
end.

