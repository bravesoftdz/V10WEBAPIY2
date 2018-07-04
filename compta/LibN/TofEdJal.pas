unit TofEdJal;

interface

uses Classes, StdCtrls, SysUtils, comctrls,
     UTof, HCtrls, Ent1,  TofMeth, HTB97,
     CritEdt,
{$IFDEF EAGLCLIENT}
  Maineagl, eQRS1,
{$ELSE}
  Fe_main, QRS1, DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
{$ENDIF}
     AGLInit,  // TheData
     Filtre,
     Extctrls; // TTimer

type
  TOF_EDITJAL = class(TOF_Meth)
  private
    Jal1, Jal2: THEdit;
    Exo, Devise : THValComboBox ;
    Etat: THValComboBox ;
    Nature : THValComboBox ;
    Date1, Date2 : THEdit;
    DateD, DateF: TDatetime;
    Num1, Num2: THEdit;
    labNumP, labA, labNat: THLabel;
    Ano, Trier, Recap: TCheckbox;
    xxWhere, xxOrderby, zRecap: THEdit;
    DevPivot: THEdit;
    BHelp: TToolbarButton97;
    procedure BHelpClick(Sender: TObject);
    procedure JalOnExit(Sender: TObject) ;
    procedure ExoOnChange(Sender: TObject) ;
    procedure DateOnExit(Sender: TObject) ;
    procedure NumOnExit(Sender: TObject) ;
    procedure EtatOnChange(Sender: TObject) ;
    procedure AnoClick(Sender: TObject) ;
    procedure TrierClick(Sender: TObject) ;
    procedure RecapClick(Sender: TObject) ;
    procedure InitCritereAvance ;
  public
    procedure OnArgument(Arguments : string) ; override ;
    procedure OnLoad ; override ;
    procedure OnUpdate ; override ;
    procedure OnNew ; override ;
    procedure FTimerTimer(Sender: TObject); // GCO
  {$IFDEF COMPTA}
    procedure ChargementCritEdt ; override ;
  {$ENDIF}

   procedure MySelectFiltre; override; // GCO - 16/07/2007 - FQ 19503

  end ;

implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  HEnt1,
  LicUtil,
  uLibExercice; // CExerciceVersRelatif

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 16/07/2004
Modifié le ... :   /  /
Description .. : Ajout de la gestion des filtres Relatifs
Mots clefs ... :
*****************************************************************}
procedure TOF_EDITJAL.ExoOnChange(Sender: TObject) ;
begin
  CExoRelatifToDates ( Exo.Value, Date1, Date2 );
  DateD := StrToDate(Date1.Text);
  DateF := StrToDate(Date2.Text);
end;

procedure TOF_EDITJAL.JalOnExit(Sender: TObject) ;
BEGIN
DoJalOnExit(THEdit(Sender), Jal1, Jal2);
END;

procedure TOF_EDITJAL.DateOnExit(Sender: TObject) ;
BEGIN
DoDateOnExit(THEdit(Sender), Date1, Date2, DateD, DateF);
END;

procedure TOF_EDITJAL.NumOnExit(Sender: TObject) ;
BEGIN
DoNumOnExit(THEdit(Sender), Num1, Num2);
END;

procedure TOF_EDITJAL.OnArgument(Arguments : string) ;
BEGIN
{$IFDEF EAGLCLIENT}
	TFQRS1(Ecran).FNomFiltre	:=	GetLeNom(TFQRS1(Ecran).Name); // Correction eAGL pour filtres a oter en 5.4.5
  ChargeFiltre(TFQRS1(Ecran).FNomFiltre,THComboBox(GetControl('FFILTRES')),TPageControl(GetControl('PAGES')));
{$ELSE}
{$ENDIF}
inherited ;
with TFQRS1(Ecran) do begin
PageOrder(Pages);
NatureEtat := Arguments;
{$IFDEF EAGLCLIENT}
//RRO le 14032003 - C'est accessible par tous, pas uniquement par PCL.
//if (ctxPCL in V_PGI.PGIContexte) and (CodeEtat='JDP') then ChoixEtat:=TRUE ;
ChoixEtat:=TRUE ;
{$ELSE}
if (ctxPCL in V_PGI.PGIContexte) and (CodeEtat='JDP') then ChoixEtat:=TRUE ;
{$ENDIF}
Etat:=THValComboBox(GetControl('FETAT')) ;
Exo:=THValComboBox(GetControl('E_EXERCICE')) ;
CInitComboExercice(Exo);
Jal1:=THEdit(GetControl('E_JOURNAL')) ;
Jal2:=THEdit(GetControl('E_JOURNAL_')) ;
Date1:=THEdit(GetControl('E_DATECOMPTABLE')) ;
Date2:=THEdit(GetControl('E_DATECOMPTABLE_')) ;
Devise:=THValComboBox(GetControl('E_DEVISE')) ;
Num1:=THEdit(GetControl('E_NUMEROPIECE')) ;
Num2:=THEdit(GetControl('E_NUMEROPIECE_')) ;
Nature:=THValComboBox(GetControl('E_NATUREPIECE')) ;
labNumP:=THLabel(GetControl('LBNUMPIECE')) ;
labA:=THLabel(GetControl('LBA')) ;
labNat:=THLabel(GetControl('LBNATURE')) ;
Ano:=TCheckBox(GetControl('ANO'));
Trier:=TCheckBox(GetControl('TRIER'));
Recap:=TCheckBox(GetControl('RECAP'));
xxWhere:=THEdit(GetControl('XX_WHERE'));
xxOrderby :=THEdit(GetControl('XX_ORDERBY'));
zRecap:=THEdit(GetControl('ZRECAP'));
DevPivot:=THEdit(GetControl('DEVPIVOT'));
BHelp:=TToolbarButton97(GetControl('BAIDE')) ;
end;
if (BHelp <> nil ) and (not Assigned(BHelp.OnClick)) then BHelp.OnClick:=BHelpClick;
if Exo<>nil then Exo.OnChange:=ExoOnChange;
if Etat<>nil then Etat.OnChange:=EtatOnChange;
if Jal1<>nil then Jal1.OnExit:=JalOnExit;
if Jal2<>nil then Jal2.OnExit:=JalOnExit;
if Date1<>nil then Date1.OnExit:=DateOnExit ;
if Date2<>nil then Date2.OnExit:=DateOnExit ;
if Num1 <> nil then Num1.OnExit := NumOnExit;
if Num2 <> nil then Num2.OnExit := NumOnExit;
if Ano<>nil then Ano.OnClick:=AnoClick;
if Trier<>nil then Trier.OnClick:=TrierClick;
if Recap<>nil then Recap.OnClick:=RecapClick;
if Nature<>nil then Nature.ItemIndex := 0 ;
if Devise<>nil then Devise.ItemIndex := 0;
if (Devise<>nil) and (DevPivot<>nil) then
  DevPivot.Text:=Devise.Items[Devise.Values.IndexOf(V_PGI.DevisePivot)];
InitCritereAvance;
TFQRS1(Ecran).Pages.ActivePage:=TFQRS1(Ecran).Pages.Pages[0];

  // Ajout GCO - 16/07/2004
  if ( CtxPCl in V_PGI.PgiContexte ) and  ( VH^.CPExoRef.Code <>'' ) then
    Exo.Value := CExerciceVersRelatif(VH^.CPExoRef.Code)
  else
    Exo.Value := CExerciceVersRelatif(VH^.Entree.Code) ;

  // GCO - 01/07/2002
  FTimer.OnTimer := FTimerTimer;
END;

procedure TOF_EDITJAL.OnNew;
begin
  inherited ;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_EDITJAL.ChargementCritEdt;
begin
  inherited;
  // Récupération des critères d'impression
  if (TheData <> nil) and (TheData is ClassCritEdt) then
  begin
    Jal1.Text  := ClassCritEdt(TheData).CritEdt.Cpt1;
    Jal2.Text  := ClassCritEdt(TheData).CritEdt.Cpt2;

    Exo.Value  := CExerciceVersRelatif(ClassCritEdt(TheData).CritEdt.Exo.Code);
    Date1.Text := DateToStr(ClassCritEdt(TheData).CritEdt.Date1);
    Date2.Text := DateToStr(ClassCritEdt(TheData).CritEdt.Date2);

    //E_ETABLISSEMENT.Value := ClassCritEdt(TheData).CritEdt.Etab;
    //E_DEVISE.Value := ClassCritEdt(TheData).CritEdt.DeviseSelect;
    TheData := nil;
  end;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/09/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_EDITJAL.OnLoad ;
begin
  inherited;
  // GCO - 19/09/2006 - Gestion BOI (Norme NF)
  FDateFinEdition := StrToDate(Date2.Text);
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_EDITJAL.OnUpdate ;
BEGIN
inherited ;
  if Trim(Jal1.Text) = '' then Jal1.Text := '000';
  if Trim(Jal2.Text) = '' then Jal2.Text := 'ZZZ';
  
  // Ajout GCO - 16/07/2004
  TFQRS1(Ecran).WhereSQL := CMajRequeteExercice ( Exo.Value, TFQRS1(Ecran).WhereSQL);

  //06/12/2006 YMO Norme NF pour journal centralisateur
  if (TFQRS1(Ecran).NatureEtat='JAC') And (FProvisoire.Text='') then  
     { BVE 29.08.07 : Mise en place d'un nouveau tracage }
{$IFNDEF CERTIFNF}
     CPEnregistreLog('JALCENTRAL '+Date1.Text+' - '+Date2.Text);
{$ELSE}
     CPEnregistreJalEvent('CEJ','Edition journal centralisateur','JALCENTRAL '+Date1.Text+' - '+Date2.Text)
{$ENDIF}
END;

procedure TOF_EDITJAL.EtatOnChange(Sender: TObject);
var fm: TFQRS1;
begin
  fm := TFQRS1(Ecran);
  if (fm.CodeEtat = 'JPN') and (Nature<>nil) then
  begin
    ControlVisible(labNat, true);
    ControlVisible(Nature, true);
  end
  else
    if (Nature<>nil) then
    begin
     ControlVisible(labNat, false);
      ControlVisible(Nature, false);
    end;

  if fm.CodeEtat = 'JCJ' then
  begin
    fm.HelpContext:=7411000;
    fm.Caption := TraduireMemoire('Journal centralisateur par journal');
  end
  else
  if (fm.CodeEtat = 'JCD') then
  begin
    fm.HelpContext:=7411000; // 14371
    fm.Caption := TraduireMemoire('Journal centralisateur par période');
  end
  else
  if fm.CodeEtat = 'JDD' then
  begin
    fm.HelpContext:=7410000;
    fm.Caption := TraduireMemoire('Journal divisionnaire par période');
  end
  else
  if fm.CodeEtat = 'JDE' then
  begin
    fm.HelpContext:=7410000;
    fm.Caption := TraduireMemoire('Journal divisionnaire par période paysage');
  end
  else
  if fm.CodeEtat = 'JDL' then
  begin
    fm.HelpContext:=7410000;
    fm.Caption := TraduireMemoire('Journal divisionnaire paysage');
  end
  else
  if fm.CodeEtat = 'JDP' then
  begin
    fm.HelpContext:=7410000;
    fm.Caption := TraduireMemoire('Journal divisionnaire');
  end
  else
  if fm.CodeEtat = 'JGE' then
  begin
  //  fm.HelpContext:=7423000;
    fm.HelpContext:=999999801;
    fm.Caption := TraduireMemoire('Journal général');
  end
  else
  if fm.CodeEtat = 'JPJ' then
  begin
    fm.HelpContext:=7413000; // 14371
    fm.Caption := TraduireMemoire('Journal périodique période/journal');
  end
  else
  if fm.CodeEtat = 'JPN' then
  begin
    fm.HelpContext:=7413000; // 14371
    fm.Caption := TraduireMemoire('Journal périodique par nature');
  end
  else
  if fm.CodeEtat = 'JPP' then
  begin
    fm.HelpContext:=7413000;
    fm.Caption := TraduireMemoire('Journal périodique journal/période');
  end;
  UpdateCaption(Ecran);
end;

procedure TOF_EDITJAL.InitCritereAvance;
var bFiltre : Boolean ;
begin
if (GetControl('FFILTRES')<>nil) and (TComboBox(GetControl('FFILTRES')).ItemIndex<>-1)
  then bFiltre:=TRUE else bFiltre:=FALSE ;
//if Ano<>nil then begin if not bFiltre then Ano.Checked:=false; AnoClick(self); end ;
if Ano<>nil then begin if not bFiltre then Ano.Checked:=true; AnoClick(self); end ;  // CA - 14/10/2003 - ANO coché par défaut
if Trier<>nil then begin if not bFiltre then Trier.Checked:=TRUE ; TrierClick(self) ; end ;
if Recap<>nil then begin if not bFiltre then Recap.Checked:=FALSE ; RecapClick(self); end ;
end;

procedure TOF_EDITJAL.AnoClick(Sender: TObject);
begin
  if Ano.Checked then xxWhere.Text:='E_QUALIFPIECE="N"'
  else xxWhere.Text:='E_QUALIFPIECE="N" AND E_ECRANOUVEAU="N"';
end;

procedure TOF_EDITJAL.TrierClick(Sender: TObject);
begin
  if Trier.Checked then xxOrderBy.Text:='E_Exercice,E_Journal,E_DateComptable,E_NumeroPiece,E_NumLigne'
  else xxOrderBy.Text:='E_Exercice,E_Journal,E_Periode,E_NumeroPiece,E_NumLigne';
end;

procedure TOF_EDITJAL.RecapClick(Sender: TObject);
begin
  if Recap.Checked then zRecap.Text:='X'
  else zRecap.Text:='';
end;

procedure TOF_EDITJAL.BHelpClick(Sender: TObject);
begin
  CallHelpTopic(Ecran) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/06/2002
Modifié le ... : 29/11/2004
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_EDITJAL.FTimerTimer(Sender: TObject);
begin
  if FCritEdtChaine <> nil then
  begin
    with FCritEdtChaine do
    begin
      if CritEdtChaine.UtiliseCritStd then
      begin
        Exo.Value := CritEdtChaine.Exercice.Code ;
        Date1.Text := DateToStr(CritEdtChaine.Exercice.Deb);
        Date2.Text := DateToStr(CritEdtChaine.Exercice.Fin);
      end;
    end;
  end;
  inherited;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 16/07/2007
Modifié le ... :   /  /    
Description .. : FQ 19503
Mots clefs ... : 
*****************************************************************}
procedure TOF_EDITJAL.MySelectFiltre;
var lExoDate : TExoDate;
begin
  inherited;
  lExoDate := CtxExercice.QuelExoDate(CRelatifVersExercice(Exo.Value));

  if (CtxExercice.QuelExo(Date1.Text, False) <> lExoDate.Code) and
     (CtxExercice.QuelExo(Date2.Text, False) <> lExoDate.Code) then
  begin
    CExoRelatifToDates(Exo.Value, Date1, Date2);
  end;
end;
////////////////////////////////////////////////////////////////////////////////

initialization
RegisterClasses([TOF_EDITJAL]) ;

end.

