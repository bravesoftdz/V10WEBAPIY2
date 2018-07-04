{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 11/02/2003
Modifié le ... :   /  /
Description .. : Passage en eAGL
Mots clefs ... :
*****************************************************************}
unit TofEdCpt;

interface
uses Classes, StdCtrls, SysUtils, Spin, comctrls,
     UTof, HCtrls, ParamSoc,
{$IFDEF EAGLCLIENT}
  Maineagl, eQRS1,uTOB,
{$ELSE}
  Fe_main, QRS1, DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
{$ENDIF}
     Filtre,
     Ent1, TofMeth, HTB97;

procedure CCLanceFiche_EPTIERSP(psz : String);

type
  TOF_EDITCPT = class(TOF_Meth)
  private
    CptG1, CptG2: THEdit;
    Exo, Devise, Etat: THValComboBox ;
    Date1, Date2 : THEdit;
    DateD, DateF: TDatetime;
    Periode, NbPeriode, DevIdx: THEdit;
    labCpt: THLabel;
    Ruptures: TSpinEdit;
    Blocnote: TCheckbox;
    Rupture, AvecBlocnote: THEdit;
    DevPivot: THEdit;
    BHelp: TToolbarButton97;
    NatureGene: THValComboBox;
    procedure BHelpClick(Sender: TObject);
    procedure CompteOnExit(Sender: TObject) ;
    procedure ExoOnChange(Sender: TObject) ;
    procedure DateOnExit(Sender: TObject) ;
    procedure EtatOnChange(Sender: TObject) ;
    procedure InitCritereAvance ;
    procedure AuxiElipsisClick ( Sender : TObject );
  public
    procedure OnArgument(Arguments : string) ; override ;
    procedure OnNew ; override ;
    procedure OnLoad ; override ;
  end ;

  TOF_EPGLIVRE = class(TOF_Meth)
  private
    CptG1, CptG2: THEdit;
    Exo, Devise, Etat: THValComboBox ;
    Date1, Date2 : THEdit;
    DateD, DateF: TDatetime;
    Num1, Num2: THEdit;
    labCpt: THLabel;
    Ruptures: TSpinEdit;
    Blocnote, Progressif, Detail: TCheckbox;
    Rupture, xxOrderBy, AvecBlocnote, SoldeProgressif : THEdit;
    DateClotureExo, DateCloture , DateExoV8: THEdit;
    DevPivot: THEdit;
    BHelp: TToolbarButton97;
    procedure BHelpClick(Sender: TObject);
    procedure CompteOnExit(Sender: TObject) ;
    procedure ExoOnChange(Sender: TObject) ;
    procedure DateOnExit(Sender: TObject) ;
    procedure NumOnExit(Sender: TObject) ;
{$IFDEF EAGLCLIENT}
{$ELSE}
    procedure EtatOnChange(Sender: TObject) ;
{$ENDIF}
    procedure InitCritereAvance ;
  public
    procedure OnArgument(Arguments : string) ; override ;
    procedure OnNew ; override ;
    procedure OnLoad ; override ;
  end ;


  TOF_TIERSP = class(TOF_Meth)
  private
    CptA1, CptA2: THEdit;
    Nature : THValComboBox ;
    BHelp: TToolbarButton97;
    procedure CompteOnExit(Sender: TObject) ;
    procedure BHelpClick(Sender: TObject);
  public
    procedure OnArgument(Arguments : string) ; override ;
    procedure OnUpdate ; override ;
  end ;


  TOF_ETASYN = class(TOF_Meth)
  private
    Exo, Dev, Etat : THValComboBox ;
    Coll : TCheckBox;
    Ex, Ex1, NMoisN, NMoisN1, Col: THEdit;
    BHelp: TToolbarButton97;
    AuDate : THedit ;
    procedure BHelpClick(Sender: TObject);
    procedure ExoOnChange(Sender: TObject) ;
    procedure CheckBoxOnClick(Sender: TObject) ;
    procedure EtatOnChange(Sender: TObject) ;
    Procedure Adaptation_Pays ;
    Procedure AuDateExit ( Sender : TObject ) ;
  public
    procedure OnArgument(Arguments : string) ; override ;
    procedure OnNew ; override ;
    Procedure OnUpdate ; override ;
  end ;


implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcGen,
  {$ENDIF MODENT1}
  HEnt1,
  LicUtil,
  ULibExercice,
  UtilPGI,
  uLibWindows,     // TestJoker
  UTofMulParamGen;

procedure CCLanceFiche_EPTIERSP(psz : String);
begin
  AGLLanceFiche('CP', 'EPTIERSP', '', psz, 'TPF');
end;

//implementation TOF_EDITCPT
procedure TOF_EDITCPT.CompteOnExit(Sender: TObject) ;
BEGIN
if (TFQRS1(Ecran).CodeEtat = 'BVE') then begin
  if (THEdit(Sender)=CptG1) and (CptG1.Text<>'') and (CptG1.Text<'60000000') then
    CptG1.Text:='60000000';
  if (THEdit(Sender)=CptG2) and (CptG2.Text<>'') and (CptG2.Text>'7ZZZZZZZ') then
    CptG2.Text:='7ZZZZZZZ';
end;
if (THEdit(Sender)=CptG1) or (THEdit(Sender)=CptG2) then
  DoCompteOnExit(THEdit(Sender), CptG1, CptG2);
END;

procedure TOF_EDITCPT.ExoOnChange(Sender: TObject) ;
var ExoDate: TExoDate;
    MM,AA,NbMois: word;
BEGIN
  DoExoToDateOnChange(Exo, Date1, Date2);
  DateD := StrToDate(Date1.Text);
  DateF := StrToDate(Date2.Text);
  //que pour l'état BVE - Balance ventilée
  ExoDate.Deb := StrToDate(Date1.Text);
  ExoDate.Fin := StrToDate(Date2.Text);
  NOMBREPEREXO(ExoDate,MM,AA,NbMois) ;
  if NbMois > 12 then
    Date2.Text := DateToStr(FinDeMois(PlusMois(ExoDate.Deb, 11)));
END;

procedure TOF_EDITCPT.DateOnExit(Sender: TObject) ;
BEGIN
DoDateOnExit(THEdit(Sender), Date1, Date2, DateD, DateF);
END;

procedure TOF_EDITCPT.OnArgument(Arguments : string) ;
BEGIN
inherited ;
	with TFQRS1(Ecran) do begin
  PageOrder(Pages);
  NatureEtat := Arguments;
  HelpContext:=7426000;
{$IFDEF EAGLCLIENT}
  //RRO le 14032003
  ChoixEtat:=TRUE ;
{$ELSE}
{$ENDIF}
  If TFQRS1(Ecran).NatureEtat='BAU' then HelpContext:=999999806 ;
  If TFQRS1(Ecran).NatureEtat='BVE' then HelpContext:=999999805 ;
  Exo:=THValComboBox(GetControl('E_EXERCICE')) ;
  CptG1:=THEdit(GetControl('E_GENERAL')) ;
  CptG2:=THEdit(GetControl('E_GENERAL_')) ;
  Date1:=THEdit(GetControl('E_DATECOMPTABLE')) ;
  Date2:=THEdit(GetControl('E_DATECOMPTABLE_')) ;
  Devise:=THValComboBox(GetControl('E_DEVISE')) ;
  labCpt:=THLabel(GetControl('TE_GENERAL')) ;
  Etat:=THValComboBox(GetControl('FETAT')) ;
  Periode:=THEdit(GetControl('PERIODE')) ;
  NbPeriode:=THEdit(GetControl('NbPERIODE')) ;
  DevIdx:=THEdit(GetControl('DEVIDX')) ;
  Ruptures:=TSpinEdit(GetControl('RUPTURES'));
  Blocnote:=TCheckbox(GetControl('Blocnote'));
  Rupture:=THEdit(GetControl('Rupture'));
  AvecBlocnote:=THEdit(GetControl('AvecBlocnote'));
  DevPivot:=THEdit(GetControl('DEVPIVOT'));
  BHelp:=TToolbarButton97(GetControl('BAIDE')) ;
  NatureGene:=THValComboBox(GetControl('G_NATUREGENE')) ;
end;
  SetControlVisible('KILO',TFQRS1(Ecran).NatureEtat='BVE');
  SetControlVisible('PROG',TFQRS1(Ecran).NatureEtat='BVE');
  SetControlVisible('LRUPTURES',not (TFQRS1(Ecran).NatureEtat='BVE'));
  SetControlVisible('RUPTURES',not (TFQRS1(Ecran).NatureEtat='BVE'));
  SetControlVisible('BLOCNOTE',not (TFQRS1(Ecran).NatureEtat='BVE'));
  if (TFQRS1(Ecran).NatureEtat='BVE') then
    SetControlProperty('KILO','caption',TraduireMemoire('Montants en')+' k'+GetColonneSQL('DEVISE','D_SYMBOLE','D_DEVISE="'+GetParamsoc('SO_DEVISEPRINC')+'"'));
  if (BHelp <> nil ) and (not Assigned(BHelp.OnClick)) then BHelp.OnClick:=BHelpClick;
  if Exo<>nil then begin
    Exo.OnChange:=ExoOnChange ;
    Exo.Value:=VH^.Entree.Code ;
  end ;
  if CptG1<>nil then CptG1.OnExit:=CompteOnExit ;
  if CptG2<>nil then CptG2.OnExit:=CompteOnExit ;
  if Date1<>nil then Date1.OnExit:=DateOnExit ;
  if Date2<>nil then Date2.OnExit:=DateOnExit ;
  if (ComboEtab<>nil) and (ComboEtab.ItemIndex=-1) then ComboEtab.ItemIndex := 0 ;
  if Devise<>nil then Devise.ItemIndex := 0;
  if Etat<>nil then Etat.OnChange := EtatOnChange;
  if (Devise<>nil) and (DevPivot<>nil) then
    DevPivot.Text:=Devise.Items[Devise.Values.IndexOf(V_PGI.DevisePivot)];
  if NatureGene<>nil then
    begin
    if TFQRS1(Ecran).NatureEtat='BGE' then NatureGene.Datatype:='TTNATGENE' ;
    NatureGene.ItemIndex:=0;
    end ;
  InitCritereAvance;
  TFQRS1(Ecran).Pages.ActivePage:=TFQRS1(Ecran).Pages.Pages[0];

  if TFQRS1(Ecran).NatureEtat='BAU' then
  begin
    if GetParamSocSecur('SO_CPMULTIERS', false) then
    begin
      CptG1.OnElipsisClick:=AuxiElipsisClick;
      CptG2.OnElipsisClick:=AuxiElipsisClick;
    end;
  end;

END ;

procedure TOF_EDITCPT.OnNew ;
BEGIN
inherited ;
with TFQRS1(Ecran) do begin
  if CodeEtat = 'BGE' then begin
    Caption := TraduireMemoire('Balance générale');
    if (CptG1<>nil) then CptG1.Text := '';
    if (CptG2<>nil) then CptG2.Text := '';
    TypeChampCompte(labCpt, CptG1, CptG2, true);
  end;
  if CodeEtat = 'BAU' then begin
    Caption := TraduireMemoire('Balance auxiliaire');
    if (CptG1<>nil) then CptG1.Text := '';
    if (CptG2<>nil) then CptG2.Text := '';
    TypeChampCompte(labCpt, CptG1, CptG2, false);
  end;
  if CodeEtat = 'BVE' then begin
    Caption := TraduireMemoire('Balance ventilée');
    if (CptG1<>nil) then CptG1.Text := '60000000';
    if (CptG2<>nil) then CptG2.Text := '7ZZZZZZZ';
    TypeChampCompte(labCpt, CptG1, CptG2, true);
  end;
UpdateCaption(Ecran);
end;
END;


procedure TOF_EDITCPT.EtatOnChange(Sender: TObject);
begin
  OnNew;
end;

procedure TOF_EDITCPT.InitCritereAvance;
begin
  Ruptures.Value := 0;
  Rupture.Text := InttoStr(Ruptures.Value);
  Blocnote.Checked := false;
  AvecBlocnote.Text := '';
end;


{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 12/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_EDITCPT.AuxiElipsisClick( Sender : TObject );
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;

procedure TOF_EDITCPT.BHelpClick(Sender: TObject);
begin
  CallHelpTopic(Ecran) ;
end;

//fin d'implementation TOF_EDITCPT

//implementation TOF_EPGLIVRE
procedure TOF_EPGLIVRE.CompteOnExit(Sender: TObject) ;
BEGIN
if (THEdit(Sender)=CptG1) or (THEdit(Sender)=CptG2) then
  DoCompteOnExit(THEdit(Sender), CptG1, CptG2);
END;

procedure TOF_EPGLIVRE.ExoOnChange(Sender: TObject) ;
BEGIN
  DoExoToDateOnChange(Exo, Date1, Date2);
  DateD := StrToDate(Date1.Text);
  DateF := StrToDate(Date2.Text);
END;

procedure TOF_EPGLIVRE.DateOnExit(Sender: TObject) ;
BEGIN
DoDateOnExit(THEdit(Sender), Date1, Date2, DateD, DateF);
END;

procedure TOF_EPGLIVRE.NumOnExit(Sender: TObject) ;
BEGIN
DoNumOnExit(THEdit(Sender), Num1, Num2);
END;

procedure TOF_EPGLIVRE.OnArgument(Arguments : string) ;
BEGIN
inherited ;
{$IFDEF EAGLCLIENT}
	TFQRS1(Ecran).FNomFiltre	:=	GetLeNom(TFQRS1(Ecran).Name); // Correction eAGL pour filtres a oter en 5.4.5
  ChargeFiltre(TFQRS1(Ecran).FNomFiltre,THComboBox(GetControl('FFILTRES')),TPageControl(GetControl('PAGES')));
{$ELSE}
{$ENDIF}
with TFQRS1(Ecran) do begin
  PageOrder(Pages);
  NatureEtat := Arguments;
{$IFDEF EAGLCLIENT}
  //RRO le 14032003
  ChoixEtat:=TRUE ;
{$ELSE}
{$ENDIF}
  Etat:=THValComboBox(GetControl('FETAT')) ;
  Exo:=THValComboBox(GetControl('E_EXERCICE')) ;
  CptG1:=THEdit(GetControl('E_GENERAL')) ;
  CptG2:=THEdit(GetControl('E_GENERAL_')) ;
  Date1:=THEdit(GetControl('E_DATECOMPTABLE')) ;
  Date2:=THEdit(GetControl('E_DATECOMPTABLE_')) ;
  Devise:=THValComboBox(GetControl('E_DEVISE')) ;
  Num1:=THEdit(GetControl('E_NUMEROPIECE')) ;
  Num2:=THEdit(GetControl('E_NUMEROPIECE_')) ;
  labCpt:=THLabel(GetControl('TE_GENERAL')) ;
  Ruptures:=TSpinEdit(GetControl('RUPTURES'));
  Blocnote:=TCheckbox(GetControl('Blocnote'));
  Progressif:=TCheckbox(GetControl('Progressif'));
  Detail:=TCheckbox(GetControl('DETAIL'));
  Rupture:=THEdit(GetControl('Rupture'));
  xxOrderBy:=THEdit(GetControl('XX_ORDERBY'));
  AvecBlocnote:=THEdit(GetControl('AvecBlocnote'));
  SoldeProgressif:=THEdit(GetControl('SoldeProgressif'));
  DevPivot:=THEdit(GetControl('DEVPIVOT'));
  BHelp:=TToolbarButton97(GetControl('BAIDE')) ;
  DateCloture := THEdit(GetControl('DATECLOTURE'));
  DateClotureExo := THEdit(GetControl('DATECLOTUREEXO'));
  DateExoV8 := THEdit(GetControl('DATEEXOV8'));
end;
  if (BHelp <> nil ) and (not Assigned(BHelp.OnClick)) then BHelp.OnClick:=BHelpClick;
{$IFDEF EAGLCLIENT}
{$ELSE}
  if Etat<>nil then Etat.OnChange := EtatOnChange;  
{$ENDIF}
  if Exo<>nil then begin
    Exo.OnChange:=ExoOnChange ;
    Exo.Value:=VH^.Entree.Code ;
  end ;
  if CptG1<>nil then CptG1.OnExit:=CompteOnExit ;
  if CptG2<>nil then CptG2.OnExit:=CompteOnExit ;
  if Date1<>nil then Date1.OnExit:=DateOnExit ;
  if Date2<>nil then Date2.OnExit:=DateOnExit ;
  if Num1 <> nil then Num1.OnExit := NumOnExit;
  if Num2 <> nil then Num2.OnExit := NumOnExit;
  if (ComboEtab<>nil) and (ComboEtab.ItemIndex=-1) then ComboEtab.ItemIndex := 0 ;
  if Devise<>nil then Devise.ItemIndex := 0;
  if (Devise<>nil) and (DevPivot<>nil) then
    DevPivot.Text:=Devise.Items[Devise.Values.IndexOf(V_PGI.DevisePivot)];
  InitCritereAvance;
  TFQRS1(Ecran).Pages.ActivePage:=TFQRS1(Ecran).Pages.Pages[0];
END ;

procedure TOF_EPGLIVRE.InitCritereAvance ;
BEGIN
  Ruptures.Value := 0;
  Rupture.Text := InttoStr(Ruptures.Value);
  Blocnote.Checked := false;
  AvecBlocnote.Text := '';
  Progressif.Checked := true;
  SoldeProgressif.Text := 'X';
  if TFQRS1(Ecran).CodeEtat = 'GLG' then
    Detail.Checked := true
  else Detail.Enabled := false;
END;

procedure TOF_EPGLIVRE.OnNew ;
BEGIN
inherited ;
with TFQRS1(Ecran) do begin
{$IFDEF EAGLCLIENT}
//if (CodeEtat = 'GLG') then begin    Pourquoi c'était laissé ?
if (CodeEtat = 'GLG') or (CodeEtat = 'GLP') then begin
{$ELSE}
if (CodeEtat = 'GLG') or (CodeEtat = 'GLP') then begin
{$ENDIF}
  HelpContext:=7424000;
  Caption := TraduireMemoire('Grand-livre général');
  if (CptG1 <> nil) and (CptG2 <> nil) then
    TypeChampCompte(labCpt, CptG1, CptG2, true);
end else
{$IFDEF EAGLCLIENT}
//if (CodeEtat = 'GLA') then begin    Pourquoi c'était laissé ?
if (CodeEtat = 'GLA') or (CodeEtat = 'GLB')  then begin
{$ELSE}
if (CodeEtat = 'GLA') or (CodeEtat = 'GLB') then begin
{$ENDIF}
  HelpContext:=7425000;
  Caption := TraduireMemoire('Grand-livre auxiliaire');
  if (CptG1 <> nil) and (CptG2 <> nil) then
    TypeChampCompte(labCpt, CptG1, CptG2, false);
end;
UpdateCaption(Ecran);
end;
END;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 29/09/2004
Modifié le ... :   /  /    
Description .. : 
Suite ........ : FQ 13981 : SBO : 29/09/2004 : déplacement du code du
Suite ........ : OnUpdtae dans le OnLoad car dans le onUpdate, le QRS1 
Suite ........ : a déjà construit la requête, c'est trop tard pour init les zones 
Suite ........ : vides !
Suite ........ : GCO - 18/12/2007 - FQ 22051
Mots clefs ... : 
*****************************************************************}
procedure TOF_EPGLIVRE.OnLoad ;
var Q: TQuery;
    Lettrage: TCheckbox ;
    Where : THEdit ;
begin
  Where:=THEdit(GetControl('XX_WHERE')) ;
  if Where<>nil then Where.Text:='E_QUALIFPIECE="N"' ;
  // Grand-Livre: Lettrage
  Lettrage:=TCheckbox(GetControl('LETTRAGE'));
  if (Lettrage<>nil) and (Where<>nil) and (Lettrage.State<>cbGrayed) then
    begin
    if Lettrage.State=cbChecked then Where.Text:=Where.Text+' AND E_LETTRAGE<>""' ;
    if Lettrage.State=cbUnchecked then Where.Text:=Where.Text+' AND E_LETTRAGE=""' ;
    end ;

  // GCO - 18/12/2007 - FQ 22051
  if (CptG1 <> nil) and (not TestJoker(CptG1.Text)) then
  begin
    // comptes généraux
    if (CptG1 <> nil) and (Trim(CptG1.Text) = '') then CptG1.Text := '00000000';
    if (CptG2 <> nil) and (Trim(CptG2.Text) = '') then CptG2.Text := 'ZZZZZZZZ';
  end;
  // FIN GCO

  //pages avancés
  Rupture.Text:=IntToStr(Ruptures.Value);
  if (Blocnote<>nil) and Blocnote.Checked then AvecBlocnote.Text:='X' else AvecBlocnote.Text:='';
  if (Progressif<>nil) and Progressif.Checked then SoldeProgressif.Text:='X' else SoldeProgressif.Text:='';
  // Date de la dernière clôture
  if DateClotureExo<>nil then DateClotureExo.Text := DateToStr(VH^.Encours.Deb-1);
  // Date de lancement de la dernière clôture
  if DateCloture<>nil then
    begin
    Q := OpenSQL ('SELECT * FROM ECRITURE WHERE E_JOURNAL="'+GetParamSoc ('SO_JALOUVRE')+
                    '" AND E_EXERCICE="'+VH^.Encours.Code+'"',True,1);
    try
        if not Q.Eof
          then DateCloture.Text := DateToStr(Q.FindField('E_DATECREATION').AsDateTime)
          else DateCloture.Text := DateToStr(iDate1900) ;
    finally
      Ferme (Q);
      end;
    end;
  // Date de la dernière clôture en détail
  if DateExoV8<>nil then
    if VH^.ExoV8.Code = ''
      then DateExoV8.Text := DateToStr(iDate1900)
      else DateExoV8.Text := DateToStr(VH^.ExoV8.Deb-1);
end ;

{$IFDEF EAGLCLIENT}
{$ELSE}
procedure TOF_EPGLIVRE.EtatOnChange(Sender: TObject);
begin
  OnNew;
end;
{$ENDIF}


procedure TOF_EPGLIVRE.BHelpClick(Sender: TObject);
begin
  CallHelpTopic(Ecran) ;
end;

//fin d'implementation TOF_EPGLIVRE

//Tof de Tiers facturés
procedure TOF_TIERSP.BHelpClick(Sender: TObject);
begin
  CallHelpTopic(Ecran) ;
end;

procedure TOF_TIERSP.CompteOnExit(Sender: TObject) ;
BEGIN
if (THEdit(Sender)=CptA1) or (THEdit(Sender)=CptA2) then
  DoCompteOnExit(THEdit(Sender), CptA1, CptA2);
END;

procedure TOF_TIERSP.OnArgument(Arguments : string) ;
BEGIN
inherited ;
with TFQRS1(Ecran) do begin
  SetControlEnabled('FListe', false);
  CptA1:=THEdit(GetControl('T_AUXILIAIRE')) ;
  CptA2:=THEdit(GetControl('T_AUXILIAIRE_')) ;
  Nature:=THValComboBox(GetControl('T_NATUREAUXI')) ;
  BHelp:=TToolbarButton97(GetControl('BAIDE')) ;
  NatureEtat := Arguments;
  if CodeEtat = 'TFP' then begin
    Caption := TraduireMemoire('Tiers payeurs facturés par payeur');
    HelpContext:=7231000;
  end
  else if CodeEtat = 'TFA' then begin
    Caption := TraduireMemoire('Tiers payeurs facturés avec payeur');
    HelpContext:=7232000;
  end;
  UpdateCaption(Ecran);
end;
if CptA1<>nil then CptA1.OnExit:=CompteOnExit ;
if CptA2<>nil then CptA2.OnExit:=CompteOnExit ;
if Nature<>nil then Nature.ItemIndex := 0 ;
if (BHelp <> nil ) and (not Assigned(BHelp.OnClick)) then BHelp.OnClick:=BHelpClick;
TFQRS1(Ecran).Pages.ActivePage:=TFQRS1(Ecran).Pages.Pages[0];
END ;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... :   /  /
Modifié le ... : 18/12/2007
Description .. :
Suite ........ : GCO - 18/12/2007 - FQ 22051
Mots clefs ... :
*****************************************************************}
procedure TOF_TIERSP.OnUpdate ;
begin
  inherited ;
  // GCO - 18/12/2007 - FQ 22051
  if (CptA1 <> nil) and (not TestJoker(CptA1.Text)) then
  begin
    if (CptA1<>nil) and (Trim(CptA1.Text) = '') then CptA1.Text := '00000000';
    if (CptA2<>nil) and (Trim(CptA2.Text) = '') then CptA2.Text := 'ZZZZZZZZ';
  end;  
end;

procedure TOF_ETASYN.CheckBoxOnClick(Sender: TObject);
BEGIN
if (Col<>nil) and Coll.Checked then Col.Text:='X'
else Col.Text:='';
END;

procedure TOF_ETASYN.EtatOnChange(Sender: TObject);
begin
  OnNew;
end;

Procedure TOF_ETASYN.Adaptation_Pays ;
Var Delta,ExoBottom : integer ;
begin
 AuDate:=nil ;
 if pos(';'+TFQRS1(Ecran).CodeEtat+';',';RCP;BAF;')>0 then
    Begin
    //Adaptation des Contrôles (L'ortdre de tab est déjà corrigé en la Fenêtre....)
    //D'abord, on pointe sur le controle de la date ....
    AuDate:=THedit(GetControl('AuDate')) ;
    AuDate.OnExit:=AuDateExit ;
    //... Ensuite, on calcule  la base du contrôle exercice....
    ExoBottom:=Exo.Top+Exo.Height ;
    Delta:= ComboEtab.Top-ExoBottom ;
    //... Ensuite, on positione le contrôle AuDate....
    AuDate.Top:=ExoBottom+Delta ;
    THLabel(getControl('TAuDate')).Top:=AuDate.top+(AuDate.Height-THLabel(getControl('TAuDate')).Height) div 2 ;
    //... et on les rend visibles....
    SetControlVisible('TAuDate',TRUE) ;
    SetControlVisible('AuDate',TRUE) ;
    //... maintenant on re-possitione le contrôle de l'ETABLISSEMENT....
    ComboEtab.Top:=AuDate.Top+AuDate.Height+Delta ;
    THLabel(getControl('TETAB')).Top:=ComboEtab.Top+(ComboEtab.Height-THLabel(getControl('TETAB')).Height) div 2 ;
    End ;
end ;

Procedure TOF_ETASYN.AuDateExit ( Sender : TObject ) ;
var NExo : TExoDate ;
Begin
  if (CQuelExercice(StrtoDateTime(AuDate.Text),NExo)) and (NExo.Code<>Exo.Value) then
     Begin
     AuDate.Tag:=999 ;
     EXO.Value:=NExo.Code ;
     auDate.Tag:=0 ;
     End ;
End ;

procedure TOF_ETASYN.ExoOnChange(Sender: TObject) ;
var s: string;
    i, nExo, nEncours: integer;
    EXODates : TExoDate ;
BEGIN
// BPY le 14/09/2004 : Fiche n° 14251 => ajout d'un test sinon plantage lors de l'affectation du filtre
if (Exo.Value = '') then exit;
// Fin BPY
if (VH^.PaysLocalisation=CodeISOES) and (Assigned(AuDate)) and (AuDate.tag<>999) and (QuelDateDeExo(Exo.Value,ExoDates)) then //XVI 24/02/2005
   AuDate.Text:=datetimetostr(exodates.fin) ;
if NMoisN <> nil then NMoisN.Text := IntToStr(NbMoisDansExo(Exo.Value));
if NMoisN1 <> nil then begin
  s := ExoPrecedent(Exo.Value);
  if s = '' then NMoisN1.Text := ''
  else NMoisN1.Text := IntToStr(NbMoisDansExo(s));
end;
if Exo.Value = VH^.Suivant.Code then begin
  Ex.Text:='N+';
  Ex1.Text:='N';
  Exit;
end;
if Exo.Value = VH^.EnCours.Code then begin
  Ex.Text:='N';
  Ex1.Text:='N-';
  Exit;
end;
nExo := 0; nEncours := 0;
for i:=1 to 5 do
if VH^.ExoClo[i].Code=Exo.Value then
  nExo:=i
else if VH^.ExoClo[i].Code='' then begin
  nEncours:=i;
  Break;
end;
Ex.Text:='N';
Ex1.Text:='N-';
//if nExo = 0 then Exit;
nExo:=nEncours - nExo;
for i := 1 to nExo do begin
  Ex.Text := Ex.Text + '-';
  Ex1.Text := Ex1.Text + '-';
end;
END;

procedure TOF_ETASYN.OnArgument(Arguments : string) ;
BEGIN
inherited ;
with TFQRS1(Ecran) do begin
  SetControlEnabled('FListe', false);
  NatureEtat := Arguments;
  //RRO le 14032003 - Pour PGE pas PCL
  if not(ctxPCL in V_PGI.PGIContexte) then
    begin
    //RRO le 14032003 on rajoute RCP, SIG et BAF
    if (CodeEtat = 'RC1') or (CodeEtat = 'RP1') or (CodeEtat = 'RCP') or (CodeEtat = 'BAF') or (CodeEtat='SIG') then
      ChoixEtat := true
    else ChoixEtat := false;
  end ;
  Exo:=THValComboBox(GetControl('EXO')) ;
  Dev:=THValComboBox(GetControl('DEV')) ;
  Coll:=TCheckBox(GetControl('COLL')) ;
  Col:=THEdit(GetControl('COL')) ;
  Ex:=THEdit(GetControl('EX')) ;
  Ex1:=THEdit(GetControl('EX1')) ;
  NMoisN:=THEdit(GetControl('NMOISN')) ;
  NMoisN1:=THEdit(GetControl('NMOISN1')) ;
  Etat:=THValComboBox(GetControl('FETAT')) ;
  BHelp:=TToolbarButton97(GetControl('BAIDE')) ;
end;
Adaptation_Pays ;
if Exo<>nil then begin
  Exo.OnChange:=ExoOnChange ;
  Exo.Value:=VH^.Entree.Code ;
end ;
if Etat<>nil then Etat.OnChange := EtatOnChange;
if Coll <> nil then Coll.OnClick := CheckBoxOnClick;
if (ComboEtab<>nil) and (ComboEtab.ItemIndex=-1) then ComboEtab.ItemIndex := 0 ;
if Dev<>nil then Dev.ItemIndex := 0;
if (BHelp <> nil ) and (not Assigned(BHelp.OnClick)) then BHelp.OnClick:=BHelpClick;
TFQRS1(Ecran).Pages.ActivePage:=TFQRS1(Ecran).Pages.Pages[0];
END ;

procedure TOF_ETASYN.OnNew;
begin
inherited ;
with TFQRS1(Ecran) do begin
  if (Etat<>nil) then begin
    if (CodeEtat = 'RC1') or (CodeEtat = 'RC2') then begin
      SetElementComboEtat(Etat, TypeEtat, NatureEtat, 'RC1;RC2');
      if not Etat.Enabled then Etat.Enabled:=true;
    end;
    if (CodeEtat = 'RP1') or (CodeEtat = 'RP2') then begin
      SetElementComboEtat(Etat, TypeEtat, NatureEtat, 'RP1;RP2');
      if not Etat.Enabled then Etat.Enabled:=true;
    end;
  end;
  if CodeEtat = 'RC1' then begin
    HelpContext:=7432000;
    Caption := TraduireMemoire('Compte de résultat charges');
  end
  else if CodeEtat = 'RC2' then begin
    HelpContext:=7432000;
    Caption := TraduireMemoire('Compte de résultat 1ère partie');
  end
  else if CodeEtat = 'RP1' then begin
    HelpContext:=7433000;
    Caption := TraduireMemoire('Compte de résultat produits');
  end
  else if CodeEtat = 'RP2' then begin
    HelpContext:=7433000;
    Caption := TraduireMemoire('Compte de résultat 2ème partie');
  end
  else if CodeEtat = 'RCP' then begin
    HelpContext:=7434000;
    Caption := TraduireMemoire('Compte de résultat');
  end
  else if CodeEtat = 'BAF' then begin
    HelpContext:=7435000;
    Caption := TraduireMemoire('Bilan Actif');
  end
  else if CodeEtat = 'BPF' then begin
    HelpContext:=7452000;
    Caption := TraduireMemoire('Bilan Passif');
  end
  else if CodeEtat = 'SIG' then begin
    HelpContext:=7453000;
    Caption := TraduireMemoire('Soldes intermédiaires de gestion');
  end;
//  if (CodeEtat = 'BAF') or (CodeEtat = 'BPF') then begin
//    Coll.Visible:=true;
//    Coll.Checked:=false;
//    Col.Text:='';
//  end else begin
    Coll.Visible:=false;
    Coll.Checked:=false;
    Col.Text:='';
//  end;
  UpdateCaption(Ecran);
end;
end;

Procedure TOF_ETASYN.OnUpdate ;
var AuDate1  : TdateTime ;
    ExoDates : TExoDate ;
Begin
   if (VH^.PaysLocalisation=CodeISOES) and (Assigned(AuDate)) and (QuelDateDeExo(Exo.Value,ExoDates)) then
   begin
      SetControlText('DATEDEBUT',DateTimetoStr(ExoDates.Deb)) ;
      SetControlText('DATEFIN',AuDate.Text) ;

      SetControlText('DATEDEBUT1',stDate1900) ;
      SetControlText('DATEFIN1',stDate1900) ;

      AuDate1:=Plusmois(StrtoDateTime(AuDate.Text),-12) ;
      if (CQuelExercice(AuDate1,ExoDates)) then
      Begin
         SetControlText('DATEDEBUT1',DateTimetoStr(EXODates.Deb)) ;
         SetControlText('DATEFIN1',DateTimetoStr(AuDate1)) ;
      End ;
   End ;
End ;

procedure TOF_ETASYN.BHelpClick(Sender: TObject);
begin
  CallHelpTopic(Ecran) ;
end;


{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 29/09/2004
Modifié le ... :   /  /    
Description .. : 
Suite ........ : FQ 13981 : SBO : 29/09/2004 : déplacement du code du
Suite ........ : OnUpdtae dans le OnLoad car dans le onUpdate, le QRS1 
Suite ........ : a déjà construit la requête, c'est trop tard pour init les zones 
Suite ........ : vides !
Suite ........ : GCO - 18/12/2007 - FQ 22051
Mots clefs ... :
*****************************************************************}
procedure TOF_EDITCPT.OnLoad;
var
  ExoDate : TExoDate;
  MM,AA,NbMois: Word;
begin
  inherited;

  // GCO - 18/12/2007 - FQ 22051
  if (CptG1 <> nil) and (not TestJoker(CptG1.Text)) then
  begin
    if (TFQRS1(Ecran).CodeEtat = 'BVE') then
    begin
      if (CptG1<>nil) and (Trim(CptG1.Text) = '') then CptG1.Text := '60000000';
      if (CptG2<>nil) and (Trim(CptG2.Text) = '') then CptG2.Text := '7ZZZZZZZ';
    end
    else
    begin
      if (CptG1<>nil) and (Trim(CptG1.Text) = '') then CptG1.Text := '00000000';
      if (CptG2<>nil) and (Trim(CptG2.Text) = '') then CptG2.Text := 'ZZZZZZZZ';
    end;
  end;
  // FIN GCO  

  if DevIdx <> nil then DevIdx.Text := IntToStr(Devise.ItemIndex);
  //pages avancés
  Rupture.Text:=IntToStr(Ruptures.Value);
  if Blocnote.Checked then AvecBlocnote.Text:='X' else AvecBlocnote.Text:='';
  //que pour l'état VBE - Balance ventilée
  ExoDate.Deb := StrToDate(Date1.Text);
  ExoDate.Fin := StrToDate(Date2.Text);
  NOMBREPEREXO(ExoDate,MM,AA,NbMois) ;
  if Periode <> nil then Periode.Text := IntToStr(GetPeriode(ExoDate.Deb));
  if NbPeriode <> nil then NbPeriode.Text := IntToStr(NbMois);
  if DevIdx <> nil then DevIdx.Text := IntToStr(Devise.ItemIndex);
end;

initialization
RegisterClasses([TOF_EDITCPT]) ;
RegisterClasses([TOF_EPGLIVRE]) ;
RegisterClasses([TOF_TIERSP]) ;
RegisterClasses([TOF_ETASYN]) ;

end.
