{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 14/03/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPREPARTITIONANA ()
Mots clefs ... : TOF;CPREPARTITIONANA
*****************************************************************}
Unit CPREPARTITIONANA_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     uTob,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
{$else}
     eMul,
     MaineAGL,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     CritEdt,
     SaisUtil;

Type
  TOF_CPREPARTITIONANA = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private
      fLesSections  : TOB;
      fGeneral      : string;
      fAxe          : string;
      fListe        : THGrid;
      fCritEdt      : TCritEdt ;
      fTypeAction   : TActionFiche;
      fOnSaveKeyDownEcran: procedure(Sender: TObject; var Key: Word; Shift:
        TShiftState) of object;
      fnbVisibleOnglet : integer;
      // évènements
      procedure OnChangePage              ( Sender : TObject );
      procedure OnDblClickListe           ( Sender : TObject );
      procedure OnClickBImprimer          ( Sender : TObject );
      procedure OnClickBCherche           ( Sender : TObject );
      procedure OnListeKeyDown            ( Sender: TObject; var Key: Word; Shift: TShiftState);
      procedure OnChangeComboExercice     ( Sender : TObject );
      procedure OnChangeDate              ( Sender : TObject );
      procedure OnKeyDownEcran            ( Sender: TObject; var Key: Word; Shift: TShiftState);
      // méthodes privées
      function AfficheLesOnglets : integer;
      procedure ChargeGrilleRepartition (pstAxe, pstCompte : string; pDateDeb, pDateFin : TDateTime);
      procedure AfficheTotalisation;
      procedure RafraichirGrille;
      procedure AfficheBoutonRecherche (bOk : boolean);
      procedure OnAfterFormShow;
  end ;

procedure CPLanceFiche_RepartitionAnalytique ( Comment : TActionFiche; pstCompte : string);

Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  ULibExercice,
  {$ENDIF MODENT1}
  Windows         // VK_
  , AGLInit       // TheData
  , HTB97
  , HSysMenu
  , Vierge
  , Ent1
  , CPMULANA_TOF
  , uObjEtats;

procedure CPLanceFiche_RepartitionAnalytique ( Comment : TActionFiche; pstCompte : string);
begin
  case Comment of
    taConsult : AGLLanceFiche('CP','CPREPARTITIONANA','','','ACTION=CONSULTATION;'+pstCompte);
    taModif   : AGLLanceFiche('CP','CPREPARTITIONANA','','','ACTION=MODIFICATION;'+pstCompte);
  end ;
end;

procedure TOF_CPREPARTITIONANA.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPREPARTITIONANA.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPREPARTITIONANA.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPREPARTITIONANA.OnLoad ;
begin
  Inherited ;
  SetFocusControl('LISTE');
end ;

procedure TOF_CPREPARTITIONANA.OnArgument (S : String ) ;
var
  i : integer;
  lStArgument : string;
begin
  Inherited ;

  lStArgument := S;
  fTypeAction := StringToAction(ReadTokenSt( lStArgument ));
  fGeneral := ReadTokenSt( lStArgument );

  if (TheData <> nil) and (TheData is (ClassCritEdt)) then
    fCritEdt := ClassCritEdt(TheData).CritEdt;

  TPageControl(GetControl('PAGES')).OnChange := OnChangePage;
  // Paramétrage de la liste
  fListe := THGrid(GetControl('LISTE'));
  fListe.OnDblClick := OnDblClickListe;
  fListe.OnKeyDown := OnListeKeyDown;
//  fListe.ColWidths[0] := 20;
  fListe.ColWidths[1] := 100;
  fListe.ColAligns[1] := taCenter;
  fListe.ColWidths[2] := 300;
  fListe.ColAligns[2] := taLeftJustify;
  fListe.ColWidths[3] := 100;
  fListe.ColAligns[3] := taRightJustify;
  fListe.ColWidths[4] := 100;
  fListe.ColAligns[4] := taRightJustify;

  // Branchement des évènements
  TToolbarButton97(GetControl('BVISUALISER')).OnClick := OnDblClickListe;
  TToolbarButton97(GetControl('BIMPRIMER')).OnClick := OnClickBImprimer;
  TToolbarButton97(GetControl('BCHERCHE')).OnClick := OnClickBCherche;
  TToolbarButton97(GetControl('BCHERCHE_')).OnClick := OnClickBCherche;
  THValComboBox(GetControl('FEXERCICE')).OnChange := OnChangeComboExercice;
  THEdit(GetControl('FDATE')).OnChange := OnChangeDate;
  THEdit(GetControl('FDATE_')).OnChange := OnChangeDate;
  fOnSaveKeyDownEcran := Ecran.OnKeyDown;
  Ecran.OnKeyDown := OnKeyDownEcran;
  TFVierge(Ecran).OnAfterFormShow := OnAfterFormShow;

  // Initialisation de la liste
  fLesSections := TOB.Create ('',nil,-1);

  // Rappel des informations
  SetControlText ('LINFO1',TraduireMemoire('Compte général')+ ' : '+fGeneral);
  SetControlText ('FEXERCICE', fCritEdt.Exo.Code);
  SetControlText ('FDATE', DateToStr(fCritEdt.Date1));
  SetControlText ('FDATE_', DateToStr(fCritEdt.Date2));
  AfficheBoutonRecherche (True);

  i := AfficheLesOnglets;
  if i >= 0 then
  begin
    SetActiveTabSheet('PAXE'+IntToStr(AfficheLesOnglets));
    fAxe := 'A'+IntToStr(i);
    // RafraichirGrille;
  end else Close;
end ;

procedure TOF_CPREPARTITIONANA.OnClose ;
begin
  fLesSections.Free;
  Inherited ;
end ;

procedure TOF_CPREPARTITIONANA.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPREPARTITIONANA.OnCancel () ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 14/03/2007
Modifié le ... : 02/10/2007
Description .. : Affichage de la répartition analytique pour l'axe courant
Mots clefs ... :
*****************************************************************}
procedure TOF_CPREPARTITIONANA.ChargeGrilleRepartition(pstAxe,pstCompte : string; pDateDeb, pDateFin : TDateTime);
var stSQL : string;
begin
  // GCO - 02/10/2007 - FQ 21029 + modif SOCREF FICHE CPREPARTITIONANA (DBIndicator à TRUE)
  stSQL := 'select y_section SECTION, S_LIBELLE LIBELLE, sum(y_debit) DEBIT ,sum(y_credit) CREDIT from analytiq ';
  stSQL := StSQL + 'LEFT JOIN SECTION ON S_AXE = Y_AXE AND S_SECTION = Y_SECTION ';
  stSQL := stSQL + ' where y_general="'+pstCompte+'" and y_axe="'+pstAxe+'" ';
  stSQL := stSQL + ' and y_datecomptable>="'+USDateTime(pDateDeb)+'" ';
  stSQL := stSQL + ' and y_datecomptable<="'+USDateTime(pDateFin)+'" ';
  stSQL := stSQL + ' and y_qualifpiece<="N" ';
  stSQL := stSQL + ' group by y_section, s_libelle ';
  fListe.VidePile(True);
  fLesSections.LoadDetailFromSQL(stSQL,False,True);
  fLesSections.PutGridDetail(fListe,False,False,'SECTION;LIBELLE;DEBIT;CREDIT');
  THSystemMenu(TFVierge(ECRAN).HMTrad).ResizeGridColumns(fListe);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 14/03/2007
Modifié le ... :   /  /
Description .. : Affichage des onglets en fonctions des axes paramétrés sur
Suite ........ : le compte général
Mots clefs ... :
*****************************************************************}
function TOF_CPREPARTITIONANA.AfficheLesOnglets : integer;
var
  Q : TQuery;
  i, iActiveTabSheet : integer;
  bVisible : boolean;
begin
  fnbVisibleOnglet := 0;
  iActiveTabSheet := -1;
  Q := OpenSQL ('SELECT G_VENTILABLE1,G_VENTILABLE2,G_VENTILABLE3,G_VENTILABLE4,'+
      ' G_VENTILABLE5 FROM GENERAUX WHERE G_GENERAL="'+fGeneral+'"',True);
  try
    if not Q.Eof then
    begin
      for i:=1 to 5 do
      begin
        // Affichage de l'onglet
        bVisible := Q.FindField('G_VENTILABLE'+IntToStr(i)).AsString='X';
        SetControlVisible('PAXE'+IntToStr(i),bVisible);
        // On récupère l'indice de la Page Visible = Axe courant
        if bVisible and (iActiveTabSheet=-1) then iActiveTabSheet := i;
        // Mise à jour du libellé de l'onglet
        if bVisible then
        begin
          SetControlText ('PAXE'+IntToStr(i),RechDom('TTAXE','A'+IntToStr(i),False));
          Inc(fnbVisibleOnglet);
        end;
      end;
    end;
  finally
    Ferme (Q);
  end;
  result := iActiveTabSheet;
end;

procedure TOF_CPREPARTITIONANA.OnChangePage(Sender: TObject);
begin
  fAxe := 'A'+IntToStr(TTabSheet(GetActiveTabSheet('PAGES')).TabIndex+1);
  RafraichirGrille;
end;

procedure TOF_CPREPARTITIONANA.OnDblClickListe(Sender: TObject);
var
  ACritEdt : ClassCritEdt;
begin
  ACritEdt := ClassCritEdt.Create;
  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  ACritEdt.CritEdt.Date1 := fCritEdt.Date1;
  ACritEdt.CritEdt.Date2 := fCritEdt.Date2;

  ACritEdt.CritEdt.DateDeb := ACritEdt.CritEdt.Date1;
  ACritEdt.CritEdt.DateFin := ACritEdt.CritEdt.Date2;

  ACritEdt.CritEdt.Exo := fCritEdt.Exo;
  ACritEdt.CritEdt.SCpt1 := fGeneral;
  ACritEdt.CritEdt.Cpt1 := fListe.Cells[1,fListe.Row];
  ACritEdt.CritEdt.Bal.Axe := fAxe;
  TheData := ACritEdt;
  try
    MultiCritereAnaZoom(fTypeAction ,AcritEdt.CritEdt);
  finally
    TheData := nil;
    ACritEdt.Free;
  end;
end;

procedure TOF_CPREPARTITIONANA.OnClickBImprimer(Sender: TObject);
begin
  TObjEtats.GenereEtatGrille (fListe,ECRAN.Caption,False);
end;

procedure TOF_CPREPARTITIONANA.OnListeKeyDown(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  case Key of
    VK_F5 :
      begin
        OnDblClickListe(nil);
        Key := 0;
      end;
  end;
end;

procedure TOF_CPREPARTITIONANA.AfficheTotalisation;
var
  stTotal : string;
  TotalDebit, TotalCredit, Solde: double;
begin
  TotalDebit := Arrondi(fLesSections.Somme('DEBIT',[''],[''],False), V_PGI.OkDecV);
  TotalCredit := Arrondi(fLesSections.Somme('CREDIT',[''],[''],False), V_PGI.OkDecV);
  Solde := Arrondi (TotalDebit - TotalCredit, V_PGI.OkDecV);
  stTotal :=  'Total débit' + ' : ' + StrS(TotalDebit,V_PGI.OkDecV) + ' - ' +
              'Total crédit ' + ' : ' + StrS(TotalCredit,V_PGI.OkDecV) + ' - '+
              'Solde' + ' : ' + StrS(Solde,V_PGI.OkDecV);
  SetControlText ('FTOTAL',stTotal);
end;

procedure TOF_CPREPARTITIONANA.OnChangeComboExercice(Sender: TObject);
var
  stExo : string;
  Exo : TExoDate;
begin
  stExo := GetControlText('FEXERCICE');
  if (stExo = '') then
  begin
    SetControlText('FDATE',DateToStr(iDate1900));
    SetControlText('FDATE_',DateToStr(iDate2099));
  end else
  begin
    QuelDateDeExo(stExo,Exo) ;
    SetControlText('FDATE',DateToStr(Exo.Deb));
    SetControlText('FDATE_',DateToStr(Exo.Fin));
  end;
  AfficheBoutonRecherche (False);
end;

procedure TOF_CPREPARTITIONANA.OnChangeDate(Sender: TObject);
begin
  AfficheBoutonRecherche (False);
end;

procedure TOF_CPREPARTITIONANA.OnClickBCherche(Sender: TObject);
begin
  RafraichirGrille;
  AfficheBoutonRecherche (True);
end;

procedure TOF_CPREPARTITIONANA.RafraichirGrille;
var stDateDeb, stDateFin : string;
begin
  stDateDeb := GetControlText('FDATE');
  stDateFin := GetControlText('FDATE_');
  if (IsValidDate(stDateDeb)) and (IsValidDate(stDateFin)) then
  begin
    ChargeGrilleRepartition(fAxe,fGeneral,StrToDate(stDateDeb),StrToDate(stDateFin));
    AfficheTotalisation;
  end;
end;

procedure TOF_CPREPARTITIONANA.AfficheBoutonRecherche(bOk: boolean);
begin
  SetControlVisible('BCHERCHE_',not bOk);
  SetControlVisible('BCHERCHE',bOk);
end;

procedure TOF_CPREPARTITIONANA.OnKeyDownEcran(Sender: TObject;
  var Key: Word; Shift: TShiftState);
var
  lInIdx : integer;
begin
  case Key of
    VK_F9 :
      begin
        OnClickBCherche(nil);
        Key := 0;
      end;
    VK_TAB :  if  ( (Shift=[ssCtrl]) or (Shift=[ssShift,ssCtrl]) ) then
              begin
                lInIdx := TTabSheet(GetActiveTabSheet('PAGES')).TabIndex;
                SetActiveTabSheet('PAXE'+IntToStr(lInIdx));
                if Shift=[ssCtrl] then
                begin
                  if lInIdx < (fNbVisibleOnglet-1)
                     then lInIdx := lInIdx + 1
                     else lInIdx := 0 ;
                   end
                 else if Shift=[ssShift,ssCtrl] then
                   begin
                   if lInIdx > 0
                     then lInIdx := lInIdx - 1
                     else lInIdx := fNbVisibleOnglet - 1 ;
                   end ;
                 SetActiveTabSheet('PAXE'+IntToStr(lInIdx));
               end ;
      else
        fOnSaveKeyDownEcran(Sender, Key, Shift);
  end;
end;

procedure TOF_CPREPARTITIONANA.OnAfterFormShow;
begin
  RafraichirGrille;
  THSystemMenu(TFVierge(ECRAN).HMTrad).ResizeGridColumns(fListe);
end;

Initialization
  registerclasses ( [ TOF_CPREPARTITIONANA ] ) ;
end.
