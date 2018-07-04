{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 25/03/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPTVAACC ()
Mots clefs ... : TOF;CPTVAACC
*****************************************************************}
Unit CPTVAACC_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     uTOb,      // TQuery
     eMUL,      // TFMul
     MaineAGL,  // AGLLanceFiche
{$ELSE}
     HQry,
     Hdb,
     db,
     {$IFNDEF DBXPRESS} dbtables {$ELSE} uDbxDataSet {$ENDIF},
     MUL,       // TFMul
     FE_Main,   // AGLLanceFiche
     CritEdt,
     UtilEdt,
{$ENDIF}
     Saisie,    // TrouveEtLanceSaisie
     forms,
     sysutils,
     ComCtrls,
     ExtCtrls,  // TRadioGroup
     HCtrls,
     HEnt1,
     HMsgBox,
     utilPGI,   // _Blocage
     HTB97,     // TToolBarButton97
     Ent1,      // ExoToDates, PositionneEtabUser
     HPanel,    // THPanel
     SaisUtil,  // RMVT, TOBM
     SaisComm,  // WhereEcriture
     TvaLettr,  // TvaAccLettrage
     UTOF ;

Procedure TvaModifAcc;

Type
  TOF_CPTVAACC = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure InitMsgBox;

    procedure E_EXERCICEChange(Sender: TObject);
    procedure BZoomPieceClick(Sender: TObject);
    procedure RAccClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure BDetailFactClick(Sender: TObject);
    procedure BImpListeClick(Sender: TObject);
  private
    HM: THMsgBox;
    FEcran    : TFMul ;
end ;

Implementation

uses
  {$IFDEF MODENT1}
  ULibExercice,
  {$ENDIF MODENT1}
  {$IFDEF eAGLCLIENT}
  MenuOLX
  {$ELSE}
  MenuOLG
  {$ENDIF eAGLCLIENT}
    ;

procedure TvaModifAcc ;
begin
  if _Blocage(['nrCloture','nrBatch','nrLettrage','nrEnca','nrDeca'],True,'nrSaisieModif') then Exit ;
  AGLLanceFiche('CP','CPTVAACC','','','');
end;

procedure TOF_CPTVAACC.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPTVAACC.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPTVAACC.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPTVAACC.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CPTVAACC.OnArgument (S : String ) ;
var
  E_JOURNAL : THValComboBox;
begin
  Inherited ;

  FEcran := TFMul(Ecran) ;
  // Création des contrôles
  HM := THMsgBox.create(FMenuG);
  InitMsgBox;

  // Evénements des contrôles
  TToolBarButton97(GetControl('BZOOMPIECE',True)).OnClick := BZoomPieceClick;
  TToolBarButton97(GetControl('BDETAILFACT',True)).OnClick := BDetailFactClick;
  TToolBarButton97(GetControl('BIMPLISTE',True)).OnClick := BImpListeClick;
  THValComboBox(GetControl('E_EXERCICE',True)).OnChange := E_EXERCICEChange;
  TRadioGroup(GetControl('RACC',True)).OnClick := RAccClick;
{$IFDEF EAGLCLIENT}
  THGrid(GetControl('FLISTE',True)).OnDblClick := FListeDblClick;
{$ELSE}
  THDBGrid(GetControl('FLISTE',True)).OnDblClick := FListeDblClick;
{$ENDIF}

  //////////////////////////////////////////////////////////////////
  // FormShow
  TFMul(Ecran).FNomFiltre := 'TVAACC';

  SetControlText('E_DEVISE',V_PGI.DevisePivot);
  if VH^.CPExoRef.Code<>'' then begin
    SetControlText('E_EXERCICE',VH^.CPExoRef.Code);
    SetControlText('E_DATECOMPTABLE',DateToStr(VH^.CPExoRef.Deb));
    SetControlText('E_DATECOMPTABLE_',DateToStr(VH^.CPExoRef.Fin));
    end
  else begin
    SetControlText('E_EXERCICE',VH^.Entree.Code);
    SetControlText('E_DATECOMPTABLE',DateToStr(V_PGI.DateEntree));
    SetControlText('E_DATECOMPTABLE_',DateToStr(V_PGI.DateEntree));
  end;
  SetControlText('E_DATEECHEANCE',StDate1900);
  SetControlText('E_DATEECHEANCE_',StDate2099);
  SetControlText('E_DATECREATION',StDate1900);
  SetControlText('E_DATECREATION_',StDate2099);
  PositionneEtabUser(THValComboBox(GetControl('E_ETABLISSEMENT',True)));

  E_JOURNAL := THValComboBox(GetControl('E_JOURNAL',True));
  if ((E_JOURNAL.Value='') and (E_JOURNAL.Values.Count>0)) then begin
    if Not E_JOURNAL.Vide then
      E_JOURNAL.Value:=E_JOURNAL.Values[0]
    else
      if E_JOURNAL.Values.Count>1 then E_JOURNAL.Value:=E_JOURNAL.Values[1] ;
  end;
  SetControlText('XX_WHERE3','E_DEVISE="'+V_PGI.DevisePivot+'"');
end ;

procedure TOF_CPTVAACC.OnClose ;
begin
  Inherited ;
  if TFMul(Ecran).Parent is THPanel then _Bloqueur('nrSaisieModif',False) ;  
  HM.Free;
end ;

procedure TOF_CPTVAACC.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPTVAACC.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_CPTVAACC.BDetailFactClick(Sender: TObject);
var
  M : RMVT ;
  QuEcr : TQuery ;
  OEcr : TOBM ;
{$IFDEF EAGLCLIENT}
  Q : TQuery;
{$ELSE}
  Q : THQuery;
{$ENDIF}
begin
{$IFDEF EAGLCLIENT}
  Q := FEcran.Q.TQ;
{$ELSE}
  Q := FEcran.Q;
{$ENDIF}
  if Not TrouveSaisie(Q,M,GetControlText('E_QUALIFPIECE')) then Exit ;
  OEcr:=Nil ;
  M.NumLigne:=Q.FindField('E_NUMLIGNE').AsInteger ;
  M.NumEche:=Q.FindField('E_NUMECHE').AsInteger ;
  QuEcr:=OpenSQL('SELECT * FROM ECRITURE WHERE '+WhereEcriture(tsGene,M,True),True) ;
  if Not QuEcr.EOF then begin
    OEcr:=TOBM.Create(EcrGen,'',True) ;
    OEcr.ChargeMvt(QuEcr) ;
  end;
  Ferme(QuEcr) ;
 if OEcr<>Nil then begin
   TvaAccLettrage(OEcr) ;
   OEcr.Free ;
 end;
end;

procedure TOF_CPTVAACC.BImpListeClick(Sender: TObject);
{$IFDEF EAGLCLIENT}
begin
// A FAIRE Voir QRTvaAcc.PAS
{$ELSE}
var Crit : TCritEdt ;
    D1,D2 : TDateTime ;
begin
  Fillchar(Crit,SizeOf(Crit),#0) ;
  Crit.NatureEtat:=neGL ;
  InitCritEdt(Crit) ;
  D1:=StrToDate(TEdit(GetControl('E_DATECOMPTABLE')).Text) ; D2:=StrToDate(TEdit(GetControl('E_DATECOMPTABLE_')).Text) ;
  Crit.Date1:=D1 ; Crit.Date2:=D2 ;
  Crit.DateDeb:=Crit.Date1 ; Crit.DateFin:=Crit.Date2 ;
  Crit.GL.Deductible:=TRadioGroup(GetControl('RACC',True)).ItemIndex=1 ;
//  EditionTvaAcc(Crit) ;
{$ENDIF}
end;

procedure TOF_CPTVAACC.BZoomPieceClick(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
  TrouveEtLanceSaisie(TFMul(Ecran).Q.TQ,taConsult,GetControlText('E_QUALIFPIECE'));
{$ELSE}
  TrouveEtLanceSaisie(TFMul(Ecran).Q,taConsult,GetControlText('E_QUALIFPIECE'));
{$ENDIF}
end;

procedure TOF_CPTVAACC.E_EXERCICEChange(Sender: TObject);
begin
	ExoToDates(GetControlText('E_EXERCICE'), TEdit(GetControl('E_DATECOMPTABLE')), TEdit(GetControl('E_DATECOMPTABLE_')) );
end;

procedure TOF_CPTVAACC.FListeDblClick(Sender: TObject);
begin
  BZoomPieceClick(Nil) ;
end;

procedure TOF_CPTVAACC.RAccClick(Sender: TObject);
begin
 {17/09/2007 YMO Ajout des règlements à la requête}
if TRadioGroup(GetControl('RACC',True)).ItemIndex=0 then SetControlText('XX_WHERE2','E_NATUREPIECE="OC" OR E_NATUREPIECE="RC"')
                                                    else SetControlText('XX_WHERE2','E_NATUREPIECE="OF" OR E_NATUREPIECE="RF"');

end;

procedure TOF_CPTVAACC.InitMsgBox;
begin
  HM.Mess.Add('ATTENTION. Modification non enregistrée.');
end;

Initialization
  registerclasses ( [ TOF_CPTVAACC ] ) ; 
end.
