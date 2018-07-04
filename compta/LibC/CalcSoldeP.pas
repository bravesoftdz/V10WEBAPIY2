unit CalcSoldeP;

interface

uses
  Windows
  , Messages
  , SysUtils
  , Classes
  , Graphics
  , Controls
  , Forms
  , Dialogs
  , StdCtrls
  , Hcompte
  , Ent1
  , HEnt1
  , Mask
  , Hctrls
  , HSysMenu
{$IFNDEF DBXPRESS}
  {$IFNDEF DBXPRESS},dbtables{$ELSE},uDbxDataSet{$ENDIF}
{$ELSE}
  ,uDbxDataSet
{$ENDIF}
  , Buttons
  , ExtCtrls
  , HStatus
  , hmsgbox
  , CpteSAV
  , utilPGI     // EstTablePartagee
  , HTB97
  , LookUp          // LookUpList
  , UtilSais        // CUpdateCumulsPointeMS
  ;

procedure CalcTotP;

type
  TFSoldeP = class(TForm)
    HMTrad: THSystemMenu;
    Panel1: TPanel;
    HPB: TPanel;
    BValider: THBitBtn;
    BFerme: THBitBtn;
    HMess: THMsgBox;
    TCPTEDEBUT: THLabel;
    TPP: THLabel;
    DP: THLabel;
    CP: THLabel;
    TOTDEBPP: THCritMaskEdit;
    TOTCREPP: THCritMaskEdit;
    TPD: THLabel;
    DD: THLabel;
    CD: THLabel;
    TOTDEBPD: THCritMaskEdit;
    TOTCREPD: THCritMaskEdit;
    BVoirSP: TToolbarButton97;
    CPTEBQ: THCritMaskEdit;
    procedure BValiderClick(Sender: TObject);
    procedure CPTEBQChange(Sender: TObject);
    procedure BVoirSPClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure CPTEBQElipsisClick(Sender: TObject);
  private
    { Déclarations privées }
    procedure LanceTraitement;
    procedure GereEnabled(OkOk: Boolean);
    procedure ReinitMontantP;
    procedure AfficheSP;
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  UTob;

procedure CalcTotP;
var
  TT: TFSoldeP;
begin
  TT := TFSoldeP.Create(Application);
  try
    TT.ShowModal;
  finally
    TT.Free;
  end;
end;

procedure TFSoldeP.GereEnabled(OkOk: Boolean);
begin
  DP.Enabled := OkOk;
  CP.enabled := OkOk;
  TOTDEBPP.Enabled := OkOk;
  TOTCREPP.Enabled := OkOk;
  TPP.Enabled := OkOk;
  DD.Enabled := OkOk;
  CD.enabled := OkOk;
  TOTDEBPD.Enabled := OkOk;
  TOTCREPD.Enabled := OkOk;
  TPD.Enabled := OkOk;
end;

procedure TFSoldeP.ReinitMontantP;
begin
  TOTDEBPP.Text := StrfMontant(0, 15, V_PGI.OkDecV, '', TRUE);
  TOTCREPP.Text := StrfMontant(0, 15, V_PGI.OkDecV, '', TRUE);
  TOTDEBPD.Text := StrfMontant(0, 15, V_PGI.OkDecV, '', TRUE);
  TOTCREPD.Text := StrfMontant(0, 15, V_PGI.OkDecV, '', TRUE);
end;

procedure TFSoldeP.AfficheSP;
var
  Q : TQuery;
  T : TOB;
  lStCompte : string;
begin
  if VH^.PointageJal then
    lStCompte := CTrouveContrePartie(CpteBQ.Text)
  else
    lStCompte := CpteBQ.Text;
  if not EstTablePartagee('GENERAUX') then begin
    Q := OpenSQL('SELECT G_TOTDEBPTP,G_TOTCREPTP,G_TOTDEBPTD,G_TOTCREPTD ' +
                 'FROM GENERAUX WHERE G_GENERAL = "' + lStCompte + '" ', TRUE);

    if not Q.Eof then
    begin
      TOTDEBPP.Text := StrfMontant(Q.Fields[0].AsFloat, 15, V_PGI.OkDecV, '', TRUE);
      TOTCREPP.Text := StrfMontant(Q.Fields[1].AsFloat, 15, V_PGI.OkDecV, '', TRUE);
      TOTDEBPD.Text := StrfMontant(Q.Fields[2].AsFloat, 15, V_PGI.OkDecV, '', TRUE);
      TOTCREPD.Text := StrfMontant(Q.Fields[3].AsFloat, 15, V_PGI.OkDecV, '', TRUE);
    end
    else
      ReinitMontantP;
    Ferme(Q);
  end
  {JP 31/05/07 : FQ 20483 : gestion des cumuls lors du partage des généraux}
  else begin
    T := TOB.Create('GENERAUX', nil, -1);
    try
      CChargeCumulsMS(fbGene, lStCompte, '', T);
      TOTDEBPP.Text := StrfMontant(T.GetDouble('G_TOTDEBPTP'), 15, V_PGI.OkDecV, '', True);
      TOTCREPP.Text := StrfMontant(T.GetDouble('G_TOTCREPTP'), 15, V_PGI.OkDecV, '', True);
      TOTDEBPD.Text := StrfMontant(T.GetDouble('G_TOTDEBPTD'), 15, V_PGI.OkDecV, '', True);
      TOTCREPD.Text := StrfMontant(T.GetDouble('G_TOTCREPTD'), 15, V_PGI.OkDecV, '', True);
    finally
      FreeAndNil(T);
    end;
  end;
end;

procedure TFSoldeP.LanceTraitement;
var
  Cpt: string;
begin
  Cpt := CpteBQ.text;
  RecalculTotPointeNew1(Cpt);
  AfficheSP;
  GereEnabled(TRUE);
end;

procedure TFSoldeP.BValiderClick(Sender: TObject);
var
  Q : TQuery;
begin
  {JP 05/06/07 : Si l'on a forcé les montants}
  if TPP.Enabled then begin
    {JP 05/06/07 : Déplacement de EstTablePartagee('GENERAUX') qui avait été placé un peu trop bas}
    if not EstTablePartagee('GENERAUX') then begin
      Q := OpenSQL('SELECT G_TOTDEBPTP,G_TOTCREPTP,G_TOTDEBPTD,G_TOTCREPTD ' +
                   'FROM GENERAUX WHERE G_GENERAL="' + CpteBQ.Text + '" ', FALSE);
      if not Q.Eof then begin
        Q.Edit;
        Q.Fields[0].AsFloat := Arrondi(Valeur(TOTDEBPP.Text), V_PGI.OkDecV);
        Q.Fields[1].AsFloat := Arrondi(Valeur(TOTCREPD.Text), V_PGI.OkDecV);
        Q.Fields[2].AsFloat := Arrondi(Valeur(TOTDEBPD.Text), V_PGI.OkDecV);
        Q.Fields[3].AsFloat := Arrondi(Valeur(TOTCREPD.Text), V_PGI.OkDecV);
        Q.Post;
      end;
      Ferme(Q);
    end
    else
      // MAJ table cumuls si besoin
      CUpdateCumulsPointeMS(CpteBQ.Text, Arrondi(Valeur(TOTDEBPP.Text), V_PGI.OkDecV),
                                         Arrondi(Valeur(TOTCREPD.Text), V_PGI.OkDecV),
                                         Arrondi(Valeur(TOTDEBPD.Text), V_PGI.OkDecV),
                                         Arrondi(Valeur(TOTCREPD.Text), V_PGI.OkDecV));
    GereEnabled(FALSE);
  end

  else begin
    try
      BeginTrans;
      LanceTraitement;
      CommitTrans;
    except
      RollBack;
    end;
  end;
end;

procedure TFSoldeP.CPTEBQChange(Sender: TObject);
begin
  GereEnabled(FALSE);
end;

procedure TFSoldeP.BVoirSPClick(Sender: TObject);
begin
  AfficheSP;
end;

procedure TFSoldeP.FormShow(Sender: TObject);
begin
  if VH^.PointageJal then
  begin
    TCPTEDEBUT.Caption := 'Journal';
    CPTEBQ.MaxLength := 3;
  end
  else
  begin
    TCPTEDEBUT.Caption := 'Compte de banque';
    CPTEBQ.MaxLength := VH^.CPta[fbGene].Lg;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/07/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFSoldeP.CPTEBQElipsisClick(Sender: TObject);
var lSt: string;
begin
  if VH^.PointageJal then
  begin
    lSt := 'SELECT J_JOURNAL, J_LIBELLE FROM JOURNAL LEFT JOIN GENERAUX ON J_CONTREPARTIE=G_GENERAL ' +
           'WHERE J_NATUREJAL = "BQE" AND J_CONTREPARTIE <> "" AND G_POINTABLE = "X" ORDER BY J_JOURNAL';

    LookUpList(THEdit(Sender), 'Journal', 'JOURNAL', 'J_JOURNAL', 'J_LIBELLE', '', 'J_JOURNAL', True, 0, lSt);
  end
  else
    LookUpList(THEdit(Sender), 'Compte général', 'GENERAUX', 'G_GENERAL', 'G_LIBELLE', 'G_POINTABLE = "X" AND G_NATUREGENE = "BQE"', 'G_GENERAL', True, 0);
end;

////////////////////////////////////////////////////////////////////////////////
end.

