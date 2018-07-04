{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 22/05/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPMODIFECHEMP ()
Mots clefs ... : TOF;CPMODIFECHEMP
*****************************************************************}
Unit CPMODIFECHEMP_TOF;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Windows,     // VK_F5
     Messages,    // WM_CLOSE
     UTOB,
{$IFDEF EAGLCLIENT}
     MaineAGL,    // AGlLanceFiche
     eMul,        // TFMul
{$ELSE}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     FE_Main,     // AGlLanceFiche
     Mul,         // TFMul
     HDB,         // THDBGrid
{$ENDIF}
{$IFDEF VER150}
     Variants,
{$ENDIF}
     Saisie,      // TrouveEtLanceSaisie
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,     // THMsgBox, Blocage, Bloqueur
     Menus,       // TMenuItem
     HTB97,       // TToolBarButton97
     HQry,        // THQuery
     Ent1,        // ExoToDates
     EcheMPA,     // ModifLeRIB
     EcheMPAMP,   // T_EcheMPAMP
     EcheUnit,    // T_ECHEUNIT
     SAISUTIL,    // QuelExo
     SaisComm,    // MvtToIdent
     ParamSoc,		// GetParamSocSecur YMO
     UTOF;

procedure ModifEcheMP;

Type
  TOF_CPMODIFECHEMP = Class (TOF)
    HM: THMsgBox;
    mnu_Echeance1: TMenuItem;
    mnu_Echeance2: TMenuItem;
{$IFDEF EAGLCLIENT}
    FListe : THGrid;
{$ELSE}
    FListe : THDBGrid;
{$ENDIF}
    Q : THQuery;
    procedure OnNew                   ; override;
    procedure OnDelete                ; override;
    procedure OnUpdate                ; override;
    procedure OnLoad                  ; override;
    procedure OnArgument (S : String ); override;
    procedure OnDisplay               ; override;
    procedure OnClose                 ; override;
    procedure OnCancel                ; override;
    procedure InitMsgBox;

    procedure E_AUXILIAIREChange(Sender: TObject);
    procedure E_EXERCICEChange(Sender: TObject);
    procedure CMSMPClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FListeDblClick(Sender: TObject);
    procedure FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure RIBNEWKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure ZoomRibClick(Sender: TObject);
    procedure BModifSerieClick(Sender: TObject);
    procedure mnu_Echeance1Click(Sender: TObject);
    procedure mnu_Echeance2Click(Sender: TObject);
    procedure bSelectAllClick(Sender: TObject);        {FP 07/11/2005 FQ16972}
    procedure FetchLesTous;                            {FP 15/06/2006 FQ16972 idem UTOFMULRELCPT}
    procedure AuxiElipsisClick(Sender : TObject);
  private
    LastQualif  : String;
    {b FP 07/11/2005 FQ16972 identique à UTOFMULRELCPT}
    bgFetchAll     : Boolean;     // Pour le bouton Sélection inversion et Select All
    AllSelecting   : boolean;
    bSelectAllClickParent     : TNotifyEvent ;   
    {e FP 07/11/2005 FQ16972}
    procedure ClickModifRIB;
    procedure ClickModifMPA;
    procedure ZoomSurRib;
    procedure ClickPourMS;
    procedure ModifSerieChamp;
    procedure ModifUneEche;
    procedure InitMulInteractif;
  public
    ClickSurBChercher : Boolean;
    {JP 28/10/05 : FQ 16876 : Gestion de la modification en série}
    procedure ModifEnSerie;
  end;

Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  ULibExercice,
  {$ENDIF MODENT1}

  {$IFDEF eAGLCLIENT}
  MenuOLX
  {$ELSE}
  MenuOLG
  {$ENDIF eAGLCLIENT}
  , Constantes
  , UTofMulParamGen; {13/04/07 YMO F5 sur Auxiliaire }


procedure ModifEcheMP;
begin
  if Blocage(['nrCloture','nrBatch'],True,'nrSaisieModif') then Exit;
  AGlLanceFiche('CP','CPMODIFECHEMP','','','');
  Bloqueur('nrSaisieModif',False);
end;

procedure TOF_CPMODIFECHEMP.OnNew;
begin
  Inherited;
end;

procedure TOF_CPMODIFECHEMP.OnDelete;
begin
  Inherited;
end;

procedure TOF_CPMODIFECHEMP.OnUpdate;
begin
  Inherited;
end;

procedure TOF_CPMODIFECHEMP.OnLoad;
begin
  Inherited;
end;

procedure TOF_CPMODIFECHEMP.OnArgument (S : String );
var
  E_JOURNAL : THValComboBox;
  lBouton   : TButton;            {FP 07/11/2005 FQ16972}
begin
  Inherited;
  bgFetchAll := True;             {FP 07/11/2005 FQ16972}

  // Création des contrôles
  HM := THMsgBox.Create(FMenuG);
  InitMsgBox;

  // Evénements des contrôles
  TMenuItem(GetConTrol('MNU_ECHEANCE1', True)).OnClick  := mnu_Echeance1Click;
  TMenuItem(GetConTrol('MNU_ECHEANCE2', True)).OnClick  := mnu_Echeance2Click;
  THEdit(GetControl('E_AUXILIAIRE', True)).OnChange := E_AUXILIAIREChange;
  THValComboBox(GetControl('E_EXERCICE', True)).OnChange := E_EXERCICEChange;
  TCheckBox(GetControl('CMSMP', True)).OnClick := CMSMPClick;
  TCheckBox(GetControl('CMSDATEECHE', True)).OnClick := CMSMPClick;
  TCheckBox(GetControl('CMSRIB', True)).OnClick := CMSMPClick;
  TFMul(Ecran).FListe.OnDblClick := FListeDblClick;
  TFMul(Ecran).FListe.OnKeyDown := FListeKeyDown;
  THEdit(GetControl('RIBNEW', True)).OnKeyDown := RIBNEWKeyDown;
  TToolBarButton97(GetControl('ZOOMRIB', True)).OnClick := ZoomRibClick;
  TToolBarButton97(GetControl('BMODIFSERIE', True)).OnClick := BModifSerieClick;
  // TToolBarButton97(GetControl('BSELECTALL',True)).OnClick := bSelectAllClick;      {FP 15/06/2006 FQ16972}
{$IFDEF EAGLCLIENT}
  FListe := THGrid(TFMul(Ecran).FListe);
{$ELSE}
  FListe := THDBGrid(TFMul(Ecran).FListe);
{$ENDIF}
  Q := TFMul(Ecran).Q;

  {b FP 07/11/2005 FQ16972 identique à UTOFMULRELCPT}
  // -> tout sélectionner
  lBouton := TButton(GetControl('bSelectAll',True)) ;
  if Assigned(lBouton.OnClick)
    then bSelectAllClickParent    := lBouton.OnClick
    else bSelectAllClickParent    := nil ;
  lBouton.OnClick := bSelectAllClick ;
  {e FP 07/11/2005 FQ16972}

  // FormShow
//  Ecran.HelpContext := 7584000;
  SetControlText('DATEECHENEW', DateToStr(V_PGI.DateEntree));
  SetControlText('E_QUALIFPIECE', 'N');
  LastQualif:='N';
  ClickSurBChercher:=FALSE;
  SetControlText('E_EXERCICE', VH^.Entree.Code);
  E_EXERCICEChange(Nil);
  SetControlText('E_DATECOMPTABLE', DateToStr(V_PGI.DateEntree));
  SetControlText('E_DATECOMPTABLE_', DateToStr(V_PGI.DateEntree));
  E_JOURNAL := THValComboBox(GetControl('E_JOURNAL', True));
  if ((E_JOURNAL.Value = '') and (E_JOURNAL.Values.Count>0)) then begin
    if Not E_JOURNAL.Vide then E_JOURNAL.Value := E_JOURNAL.Values[0]
                          else if E_JOURNAL.Values.Count>1 then E_JOURNAL.Value:=E_JOURNAL.Values[1];
  end;
//Q.Liste:='CPMODECHEMP1';

{b FP 15/06/2006 FQ16972}
{$IFDEF EAGLCLIENT}
  FListe.MultiSelect := True;
{$ELSE}
  FListe.MultiSelection := True;
{$ENDIF}
  SetControlVisible('BSELECTALL', True);
{e FP 15/06/2006 FQ16972}

  InitMulInteractif;

  if GetParamSocSecur('SO_CPMULTIERS', false) then
    THEdit(GetControl('E_AUXILIAIRE', true)).OnElipsisClick:=AuxiElipsisClick;

end;

procedure TOF_CPMODIFECHEMP.OnClose;
begin
  Inherited;
  HM.Free;
  VH^.MPPop.MPGenPop:='';
  VH^.MPPop.MPAuxPop:='';
  VH^.MPPop.MPJalPop:='';
  VH^.MPPop.MPExoPop:='';
  VH^.MPPop.MPNumPop:=0;
  VH^.MPPop.MPNumLPop:=0;
  VH^.MPPop.MPNumEPop:=0;
  VH^.MPPop.MPDatePop:=0;
end;

procedure TOF_CPMODIFECHEMP.OnDisplay ();
begin
  Inherited;
end;

procedure TOF_CPMODIFECHEMP.OnCancel ();
begin
  Inherited;
end;

procedure TOF_CPMODIFECHEMP.InitMsgBox;
begin
HM.Mess.Add('0;?Caption?;Cette échéance d''A-Nouveau n''est pas modifiable.;W;O;O;O;');
HM.Mess.Add('1;?Caption?;Les critères de compte ne sont pas remplis. Confirmez-vous la recherche sur tous les comptes ?;Q;YNC;N;N;');
HM.Mess.Add('2;Modification en série sur mode de paiement;La nouvelle valeur du mode de paiement n''est pas renseignée.;W;O;O;O;');
HM.Mess.Add('3;Modification en série sur RIB;La nouvelle valeur de RIB n''est pas renseignée.;W;O;O;O;');
HM.Mess.Add('4;Modification en série sur date d''échéance;La nouvelle valeur de la date d''échéance n''est pas renseignée.;W;O;O;O;');
end;

procedure TOF_CPMODIFECHEMP.BModifSerieClick(Sender: TObject);
begin
  {JP 28/10/05 : FQ 16876 : le bouton BModifSerie sert au 2 traitements et le
                 bouton Bouvrir sert à zoomer}
  if (GetCheckBoxState('CMSMP') = cbChecked) or
     (GetCheckBoxState('CMSDATEECHE') = cbChecked) or
     (GetCheckBoxState('CMSRIB') = cbChecked) then ModifEnSerie
                                              else ClickModifMPA;
end;

{JP 14/04/05 : FQ 15608 : gestion de e_tresosynchro
               Mise en forme plus proche des standards eAgl en remplaçant le query(.Post) par une TOB}
procedure TOF_CPMODIFECHEMP.ClickModifMPA;
var
  M : RMVT;
  Trouv,OkModif : boolean;
  TAN   : String3;
  RIB,Aux,{OldRib,}Gen : String;
  RJal,RExo,RQualif : String;
  RDate : TDateTime;
  RNumP,RNumL,RNumEche : Integer;
  EcheMPA,OldEcheMPA : T_EcheMPAMP;
  TIDTIC : Boolean;
  tEcriture : TOB;
begin
  if Q.EOF then Exit;

  TIDTIC:=FALSE;

  RJal     := GetField('E_JOURNAL');
  RExo     := GetField('E_EXERCICE');
  RQualif  := GetField('E_QUALIFPIECE');
  RDate    := GetField('E_DATECOMPTABLE');
  RNumP    := GetField('E_NUMEROPIECE');
  RNumL    := GetField('E_NUMLIGNE');
  RNumEche := GetField('E_NUMECHE');

  tEcriture := TOB.Create('ECRITURE', nil, -1);
  try
    tEcriture.SetString  ('E_JOURNAL'      , RJal    );
    tEcriture.SetString  ('E_EXERCICE'     , RExo    );
    tEcriture.SetDateTime('E_DATECOMPTABLE', RDate   );
    tEcriture.SetString  ('E_QUALIFPIECE'  , RQualif );
    tEcriture.SetInteger ('E_NUMEROPIECE'  , RNumP   );
    tEcriture.SetInteger ('E_NUMLIGNE'     , RNumL   );
    tEcriture.SetInteger ('E_NUMECHE'      , RNumEche);

    Trouv := tEcriture.LoadDB;

    if Trouv then begin
      M   := TobToIdent(tEcriture, True);
      RIB := tEcriture.GetString('E_RIB');
      TAN := tEcriture.GetString('E_ECRANOUVEAU');
      Aux := tEcriture.GetString('E_AUXILIAIRE');
      if Aux = '' then begin
        Gen := tEcriture.GetString('E_GENERAL');
        if (Gen <> '') and (RNumEche > 0) then TIDTIC := True
                                          else Trouv  := False;
      end;

      if TAN = 'OAN' then begin
        if (M.CodeD <> V_PGI.DevisePivot) or
           ((VH^.EXOV8.Code <> '') and (M.DateC < VH^.EXOV8.Deb)) then begin
          HM.Execute(0,'',''); // Cette échéance d''A-Nouveau n''est pas modifiable.
          Trouv := False;
        end;
      end;
    end;

    if Trouv then begin
      EcheMPA.TIDTIC        := TIDTIC;
      EcheMPA.DateEche      := tEcriture.GetDateTime('E_DATEECHEANCE');
      EcheMPA.DateComptable := tEcriture.GetDateTime('E_DATECOMPTABLE');
      EcheMPA.ModePaie      := tEcriture.GetString('E_MODEPAIE');
      EcheMPA.Jal           := tEcriture.GetString('E_JOURNAL');
      EcheMPA.NatP          := tEcriture.GetString('E_NATUREPIECE');

      if EcheMPA.TIDTIC then EcheMPA.Aux := Gen
                        else EcheMPA.Aux := Aux;
      EcheMPA.Rib := RIB;

      if Copy(EcheMPA.Rib, 1, 1) = '*' then System.Delete(EcheMPA.Rib, 1, 1);

      EcheMPA.Montant    := tEcriture.GetDouble('E_DEBIT') + tEcriture.GetDouble('E_CREDIT');
      EcheMPA.NumP       := tEcriture.GetInteger('E_NUMEROPIECE');
      EcheMPA.RefExterne := tEcriture.GetString('E_REFEXTERNE');
      EcheMPA.RefLibre   := tEcriture.GetString('E_REFLIBRE');
      EcheMPA.NumTraChq  := tEcriture.GetString('E_NUMTRAITECHQ');
      EcheMPA.CodeAcc    := tEcriture.GetString('E_CODEACCEPT');

      OldEcheMPA := EcheMPA;
      OkModif    := False;

      if ModifUneEcheanceMPAMP(EcheMPA) then begin
        if EcheMPA.ModePaie <> OldEcheMPA.ModePaie then begin
          tEcriture.SetString('E_MODEPAIE', EcheMPA.ModePaie);
          OkModif := True;
        end;
        if EcheMPA.CodeAcc <> OldEcheMPA.CodeAcc then begin
          tEcriture.SetString('E_CODEACCEPT', EcheMPA.CodeAcc);
          OkModif := True;
        end;
        if EcheMPA.DateEche <> OldEcheMPA.DateEche then begin
          tEcriture.SetDateTime('E_DATEECHEANCE', EcheMPA.DateEche);
          OkModif := True;
        end;
        if EcheMPA.Rib <> OldEcheMPA.Rib then begin
          {JP 28/10/05 : Problème chez SIC :
                         Pour éviter de se retrouver avec des rib qui n'ont qu'une étoile.
                         D'ailleurs, cela doit améliorer le traitement dans CFONB.DetailTRANS}
          if EcheMPA.Rib = '*' then EcheMPA.Rib := '';

          tEcriture.SetString('E_RIB', EcheMPA.Rib);
          OkModif := True;
        end;
        if EcheMPA.RefExterne <> OldEcheMPA.RefExterne then begin
          tEcriture.SetString('E_REFEXTERNE', EcheMPA.RefExterne);
          OkModif := True;
        end;
        if EcheMPA.RefLibre <> OldEcheMPA.RefLibre then begin
          tEcriture.SetString('E_REFLIBRE', EcheMPA.RefLibre);
          OkModif := True;
        end;
        if EcheMPA.NumTraChq <> OldEcheMPA.NumTraChq then begin
          tEcriture.SetString('E_NUMTRAITECHQ', EcheMPA.NumTraChq);
          OkModif := True;
        end;

        {JP 13/04/05 : FQ 15608 : mise à jour de E_TRESOSYNCHRO si le mode de paiement ou l'échéance ont changé}
        if (EcheMPA.ModePaie <> OldEcheMPA.ModePaie) or (EcheMPA.DateEche <> OldEcheMPA.DateEche) then
          tEcriture.SetString('E_TRESOSYNCHRO', ets_Nouveau);
        {Mise à jour des modifications}
        if OkModif then
          tEcriture.UpdateDB;
      end;

      Application.ProcessMessages;
      if OkModif then begin
        TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
        VH^.MPModifFaite := True;
      end;
      {$IFDEF EAGLCLIENT}
      Q.TQ.FindFirst(['E_JOURNAL;E_EXERCICE;E_DATECOMPTABLE;E_QUALIFPIECE;E_NUMEROPIECE;E_NUMLIGNE;E_NUMECHE'],
                  [RJal,RExo,RDate,RQualif,RNumP,RNumL,RNumEche],True);
      {$ELSE}
      Q.Locate('E_JOURNAL;E_EXERCICE;E_DATECOMPTABLE;E_QUALIFPIECE;E_NUMEROPIECE;E_NUMLIGNE;E_NUMECHE',
                  varArrayOf([RJal,RExo,RDate,RQualif,RNumP,RNumL,RNumEche]),[]);
      {$ENDIF}
    end;
  finally
    if Assigned(tEcriture) then FreeAndNil(tEcriture);
  end;
end;

procedure TOF_CPMODIFECHEMP.ClickModifRIB;
var
  M : RMVT;
  Q1 : TQuery;
  Trouv : boolean;
  TAN   : String3;
  RIB,Aux,OldRib : String;
  RJal,RExo,RQualif : String;
  RDate : TDateTime;
  RNumP,RNumL,RNumEche : Integer;
begin
  if Q.EOF then Exit;
{$IFDEF EAGLCLIENT}
  Q.TQ.Seek(FListe.Row-1);
{$ENDIF}
  RJal:=Q.FindField('E_JOURNAL').AsString;
  RExo:=QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime));
  RQualif:=LastQualif;
  RDate:=Q.FindField('E_DATECOMPTABLE').AsDateTime;
  RNumP:=Q.FindField('E_NUMEROPIECE').AsInteger;
  RNumL:=Q.FindField('E_NUMLIGNE').AsInteger;
  RNumEche:=Q.FindField('E_NUMECHE').AsInteger;

  Q1:=OpenSQL('SELECT * FROM ECRITURE WHERE E_JOURNAL="'+Q.FindField('E_JOURNAL').AsString+'"'
            +' AND E_EXERCICE="'+QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime))+'"'
            +' AND E_DATECOMPTABLE="'+USDATETIME(Q.FindField('E_DATECOMPTABLE').AsDateTime)+'"'
            +' AND E_QUALIFPIECE="'+LastQualif+'"'
            +' AND E_NUMEROPIECE='+Q.FindField('E_NUMEROPIECE').AsString
            +' AND E_NUMLIGNE='+Q.FindField('E_NUMLIGNE').AsString
            +' AND E_NUMECHE='+Q.FindField('E_NUMECHE').AsString,FALSE);
  Trouv := Not Q1.EOF;
  if Trouv then begin
    M:=MvtToIdent(Q1,fbGene,True);
    RIB:=Q1.FindField('E_RIB').AsString;
    TAN:=Q1.FindField('E_ECRANOUVEAU').AsString;
    Aux:=Q1.FindField('E_AUXILIAIRE').AsString;
    If Aux='' then Trouv:=FALSE;
    if TAN='OAN' then begin
      if (M.CodeD<>V_PGI.DevisePivot) or ((VH^.EXOV8.Code<>'') and (M.DateC<VH^.EXOV8.Deb)) then begin
        HM.Execute(0,'',''); // Cette échéance d''A-Nouveau n''est pas modifiable.
        Trouv:=FALSE;
      end;
    end;
  end;
  If Trouv then begin
    OldRIB:=RIB;
    ModifLeRIB(Rib,Aux);
    If RIB<>OldRib then begin
      Q1.Edit;
      Q1.FindField('E_RIB').AsString:=RIB;
      Q1.Post;
    end;
    Ferme(Q1);
    Application.ProcessMessages;
    TFmul(Ecran).BChercheClick(Nil);
{$IFDEF EAGLCLIENT}
    Q.TQ.FindFirst(['E_JOURNAL;E_EXERCICE;E_DATECOMPTABLE;E_QUALIFPIECE;E_NUMEROPIECE;E_NUMLIGNE;E_NUMECHE'],
                 [RJal,RExo,RDate,RQualif,RNumP,RNumL,RNumEche],True)
{$ELSE}
    Q.Locate('E_JOURNAL;E_EXERCICE;E_DATECOMPTABLE;E_QUALIFPIECE;E_NUMEROPIECE;E_NUMLIGNE;E_NUMECHE',
                 varArrayOf([RJal,RExo,RDate,RQualif,RNumP,RNumL,RNumEche]),[])
{$ENDIF}
    end
  else Ferme(Q1);
end;

procedure TOF_CPMODIFECHEMP.ClickPourMS;
var
  OkMS : Boolean;
begin
  OkMS := (GetCheckBoxState('CMSMP') = cbChecked) or (GetCheckBoxState('CMSDATEECHE') = cbChecked) or (GetCheckBoxState('CMSRIB') = cbChecked);

  SetControlVisible('BMODIFSERIE', OkMS);
  {JP 28/10/05 : FQ 16876 : le bouton BOuvrir sert maintenant à ouvrir la pièce
  SetControlVisible('BOUVRIR', not OkMS);}

  SetControlEnabled('MPNEW', GetCheckBoxState('CMSMP') = cbChecked);
  SetControlEnabled('DATEECHENEW', GetCheckBoxState('CMSDATEECHE') = cbChecked);
  SetControlEnabled('RIBNEW', GetCheckBoxState('CMSRIB') = cbChecked);
  If (GetCheckBoxState('CMSRIB') = cbUnChecked) then SetControlText('RIBNEW', '');
  If (GetCheckBoxState('CMSMP') = cbUnChecked) then  SetControlProperty('MPNEW', 'ITEMINDEX', -1);

(*FP 15/06/2006 FQ16972
{$IFDEF EAGLCLIENT}
  FListe.MultiSelect := OkMS;
  SetControlVisible('BSELECTALL', OkMS);
{$ELSE}
  FListe.MultiSelection := OkMS;
  SetControlVisible('BSELECTALL', OkMS);
{$ENDIF}
*)
end;

{b FP 07/11/2005 FQ16972 identique à UTOFMULRELCPT}
procedure TOF_CPMODIFECHEMP.bSelectAllClick(Sender: TObject);
begin
  AllSelecting := true;

  FetchLesTous;
  if Assigned(bSelectAllClickParent) then
    bSelectAllClickParent(Sender) ;

  AllSelecting := false;
end;
{e FP 07/11/2005 FQ16972}

procedure TOF_CPMODIFECHEMP.CMSMPClick(Sender: TObject);
begin
  ClickPourMS;
end;

procedure TOF_CPMODIFECHEMP.E_AUXILIAIREChange(Sender: TObject);
begin
  SetControlEnabled('CMSRIB', Trim(GetControlText('E_AUXILIAIRE'))<>'');
  ClickPourMS;
end;

procedure TOF_CPMODIFECHEMP.E_EXERCICEChange(Sender: TObject);
begin
  ExoToDates(GetControlText('E_EXERCICE'), GetControl('E_DATECOMPTABLE', True) , GetControl('E_DATECOMPTABLE_'));
end;

procedure TOF_CPMODIFECHEMP.FListeDblClick(Sender: TObject);
begin
  inherited;
  {JP 28/10/05 : FQ 16876 : Ergonomie : BOuvrir / FLiteDoubleClick lance le zooom}
  {$IFDEF EAGLCLIENT}
  Q.TQ.Seek(FListe.Row-1);  {26.06.07 YMO FQ20530}
  TrouveEtLanceSaisie(Q.TQ, taConsult, LastQualif);
  {$ELSE}
  TrouveEtLanceSaisie(Q, taConsult, LastQualif);
  {$ENDIF}
end;

procedure TOF_CPMODIFECHEMP.FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var
  Vide : Boolean;
begin
  Vide:=(Shift=[]);
  if Vide and (Key=VK_F5) then begin
    Key:=0;
    ClickModifMPA;
  end;
  if (ssCtrl in Shift) and (Key=VK_F5) then begin
    Key:=0;
    ClickModifRIB;
  end;
end;

procedure TOF_CPMODIFECHEMP.FormClose(Sender: TObject; var Action: TCloseAction);
begin

end;

procedure TOF_CPMODIFECHEMP.FormShow(Sender: TObject);
begin

end;

procedure TOF_CPMODIFECHEMP.InitMulInteractif;
var
  OkOk : Boolean;
  St : String;
begin
  SetControlText('XX_WHEREPOP', '');
  VH^.MPModifFaite:= False;
  OkOk := (VH^.MPPop.MPAuxPop<>'') Or ((VH^.MPPop.MPAuxPop='') And (VH^.MPPop.MPNumEPop>0));
  If OkOk then begin
    If VH^.MPPop.MPGenPop<>'' then SetControlText('E_GENERAL', VH^.MPPop.MPGenPop);
    If VH^.MPPop.MPAuxPop<>'' then SetControlText('E_AUXILIAIRE', VH^.MPPop.MPAuxPop);
    If VH^.MPPop.MPJalPop<>'' then SetControlText('E_JOURNAL', VH^.MPPop.MPJalPop);
    If VH^.MPPop.MPExoPop<>'' then SetControlText('E_EXERCICE', VH^.MPPop.MPExoPop);
    If VH^.MPPop.MPNumPop<>0 then begin SetControlText('E_NUMEROPIECE', IntToStr(VH^.MPPop.MPNumPop)); SetControlText('E_NUMEROPIECE_', IntToStr(VH^.MPPop.MPNumPop)); end;
    If VH^.MPPop.MPDatePop<>0 then begin SetControlText('E_DATECOMPTABLE', DateToStr(VH^.MPPop.MPDatePop)); SetControlText('E_DATECOMPTABLE_',DateToStr(VH^.MPPop.MPDatePop)); end;
    If VH^.MPPop.MPNumLPop<>0 then SetControlText('XX_WHEREPOP', GetControlText('XX_WHEREPOP') + ' AND E_NUMLIGNE='+IntToStr(VH^.MPPop.MPNumLPop)+' ');
    If VH^.MPPop.MPNumEPop<>0 then SetControlText('XX_WHEREPOP', GetControlText('XX_WHEREPOP') + ' AND E_NUMECHE='+IntToStr(VH^.MPPop.MPNumEPop)+' ');
  end;
  If Trim(GetControlText('XX_WHEREPOP'))<>'' then begin
    St := GetControlText('XX_WHEREPOP');
    System.Delete(St,1,4);
    SetControlText('XX_WHEREPOP', St);
  end;
end;

procedure TOF_CPMODIFECHEMP.mnu_Echeance1Click(Sender: TObject);
begin
  ModifUneEche;
end;

procedure TOF_CPMODIFECHEMP.mnu_Echeance2Click(Sender: TObject);
begin
  ClickModifMPA;
end;

procedure TOF_CPMODIFECHEMP.ModifSerieChamp;
var
  RJal,RExo,RQualif : String;
  What,SQL : String;
begin
  RJal := Q.FindField('E_JOURNAL').AsString;
  RExo := QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime));
  RQualif:=LastQualif;
  What:='';
  If (GetCheckBoxState('CMSMP') = cbChecked)       then What:=What+',E_MODEPAIE="'+GetControlText('MPNEW')+'" ';
  If (GetCheckBoxState('CMSDATEECHE') = cbChecked) then What:=What+',E_DATEECHEANCE="'+ USDateTime(StrToDate(GetControlText('DATEECHENEW')))+'" ';
  If (GetCheckBoxState('CMSRIB') = cbChecked)      then What:=What+',E_RIB="'+GetControlText('RIBNEW')+'" ';

  {JP 14/09/05 : FQ 15608 : je n'avais traité la mise à jour de E_TRESOSYNCHRO que lors du traitement
                 à "l'unité" (cf ClickModifMPA). Réparation de l'oubli}
  if (GetCheckBoxState('CMSDATEECHE') = cbChecked) or (GetCheckBoxState('CMSMP') = cbChecked) then
    What := What + ', E_TRESOSYNCHRO = "' + ets_Nouveau + '" ';

  If Trim(What)='' then Exit;
  {JP 16/11/07 : FQ 21847 : Gestion de E_UTILISATEUR}
  What := What + ', E_UTILISATEUR = "' + V_PGI.User + '", E_DATEMODIF = "' + UsTime(NowH) + '" ';

  What := Copy(What,2,Length(What)-1);
  SQL:='UPDATE Ecriture SET '+What
      +'where E_JOURNAL="'+Q.FindField('E_JOURNAL').AsString+'"'
      +' AND E_EXERCICE="'+QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime))+'"'
      +' AND E_DATECOMPTABLE="'+USDATETIME(Q.FindField('E_DATECOMPTABLE').AsDateTime)+'"'
      +' AND E_QUALIFPIECE="'+LastQualif+'"'
      +' AND E_NUMEROPIECE='+Q.FindField('E_NUMEROPIECE').AsString
      +' AND E_NUMLIGNE='+Q.FindField('E_NUMLIGNE').AsString
      +' AND E_NUMECHE='+Q.FindField('E_NUMECHE').AsString;
  ExecuteSQL(SQL);
end;

procedure TOF_CPMODIFECHEMP.ModifUneEche;
var
  M : RMVT;
  Q1 : TQuery;
  Trouv : boolean;
  EU    : T_ECHEUNIT;
  TAN   : String3;
  k     : integer;
  Coll  : String;
  RJal,RExo,RQualif : String;
  RDate : TDateTime;
  RNumP,RNumL,RNumEche : Integer;
begin
  if Q.EOF then Exit;
{$IFDEF EAGLCLIENT}
  Q.TQ.Seek(FListe.Row-1);
{$ENDIF}
  RJal:=Q.FindField('E_JOURNAL').AsString;
  RExo:=QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime));
  RQualif:=LastQualif;
  RDate:=Q.FindField('E_DATECOMPTABLE').AsDateTime;
  RNumP:=Q.FindField('E_NUMEROPIECE').AsInteger;
  RNumL:=Q.FindField('E_NUMLIGNE').AsInteger;
  RNumEche:=Q.FindField('E_NUMECHE').AsInteger;
  Q1:=OpenSQL('Select * from Ecriture where E_JOURNAL="'+Q.FindField('E_JOURNAL').AsString+'"'
             +' AND E_EXERCICE="'+QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime))+'"'
             +' AND E_DATECOMPTABLE="'+USDATETIME(Q.FindField('E_DATECOMPTABLE').AsDateTime)+'"'
             +' AND E_QUALIFPIECE="'+LastQualif+'"'
             +' AND E_NUMEROPIECE='+Q.FindField('E_NUMEROPIECE').AsString
             +' AND E_NUMLIGNE='+Q.FindField('E_NUMLIGNE').AsString
             +' AND E_NUMECHE='+Q.FindField('E_NUMECHE').AsString,True);
  Trouv:=Not Q1.EOF;
  if Trouv then begin
    M:=MvtToIdent(Q1,fbGene,True); FillChar(EU,Sizeof(EU),#0);
    EU.DateEche:=Q1.FindField('E_DATEECHEANCE').AsDateTime; EU.ModePaie:=Q1.FindField('E_MODEPAIE').AsString;
    EU.DebitDEV:=Q1.FindField('E_DEBITDEV').AsFloat; EU.CreditDEV:=Q1.FindField('E_CREDITDEV').AsFloat;
    EU.Debit:=Q1.FindField('E_DEBIT').AsFloat; EU.Credit:=Q1.FindField('E_CREDIT').AsFloat;
    EU.DEVISE:=Q1.FindField('E_DEVISE').AsString; EU.TauxDEV:=Q1.FindField('E_TAUXDEV').AsFloat;
    EU.DateComptable:=Q1.FindField('E_DATECOMPTABLE').AsDateTime;
    EU.DateModif:=Q1.FindField('E_DATEMODIF').AsDateTime;
    EU.ModeSaisie:=Q1.FindField('E_MODESAISIE').AsString;
    TAN:=Q1.FindField('E_ECRANOUVEAU').AsString;
    {#TVAENC}
    if VH^.OuiTvaEnc then begin
      Coll:=Q1.FindField('E_GENERAL').AsString;
      if EstCollFact(Coll) then begin
        for k:=1 to 4 do EU.TabTva[k]:=Q1.FindField('E_ECHEENC'+IntToStr(k)).AsFloat;
        EU.TabTva[5]:=Q1.FindField('E_ECHEDEBIT').AsFloat;
      end;
    end;
  end;
  Ferme(Q1);
  if Not Trouv then Exit;
  if TAN='OAN' then begin
    if (M.CodeD<>V_PGI.DevisePivot) or ((VH^.EXOV8.Code<>'') and (M.DateC<VH^.EXOV8.Deb)) then begin
      HM.Execute(0,'',''); // Cette échéance d''A-Nouveau n''est pas modifiable.
      Exit;
    end;
  end;
  if ModifUneEcheance(M,EU) then begin
    Application.ProcessMessages;
    TFMul(Ecran).BChercheClick(Nil);
{$IFDEF EAGLCLIENT}
    Q.TQ.FindFirst(['E_JOURNAL;E_EXERCICE;E_DATECOMPTABLE;E_QUALIFPIECE;E_NUMEROPIECE;E_NUMLIGNE;E_NUMECHE'],
                   [RJal,RExo,RDate,RQualif,RNumP,RNumL,RNumEche],True)
{$ELSE}
    Q.Locate('E_JOURNAL;E_EXERCICE;E_DATECOMPTABLE;E_QUALIFPIECE;E_NUMEROPIECE;E_NUMLIGNE;E_NUMECHE',
                 varArrayOf([RJal,RExo,RDate,RQualif,RNumP,RNumL,RNumEche]),[])
{$ENDIF}
  end;
  If ((VH^.MPPop.MPAuxPop<>'') Or ((VH^.MPPop.MPAuxPop='') And (VH^.MPPop.MPNumEPop>0))) then PostMessage(Ecran.Handle, WM_CLOSE, 0,0); // 10930
end;

procedure TOF_CPMODIFECHEMP.RIBNEWKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var
  Vide : Boolean;
begin
  Vide:=(Shift=[]);
  if Vide and (Key=VK_F5) then begin Key:=0; ZoomSurRIB; end;
  Key := 0;
end;

procedure TOF_CPMODIFECHEMP.ZoomRibClick(Sender: TObject);
begin
  ZoomSurRib;
end;

procedure TOF_CPMODIFECHEMP.ZoomSurRib;
var
  Rib,Aux : String;
begin
  Rib:='';
  Aux := GetControlText('E_AUXILIAIRE');
  If ModifLeRIB(Rib,Aux) then SetControlText('RIBNEW', Rib);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMODIFECHEMP.ModifEnSerie;
{---------------------------------------------------------------------------------------}
var
  i : Integer;
begin

  if (GetCheckBoxState('CMSMP') = cbChecked) And (GetControlText('MPNEW') = '') then begin
    HM.Execute(2,'','');  // La nouvelle valeur du mode de paiement n''est pas renseignée.
    TFMul(Ecran).Pages.ActivePage := TTabSheet(GetControl('PMODIFS', True));
    SetFocusControl('MPNEW');
    Exit;
  end;

  if (GetCheckBoxState('CMSDATEECHE') = cbChecked) And (Not IsValidDate(GetControlText('DATEECHENEW'))) then begin
    HM.Execute(4,'','');  // La nouvelle valeur de la date d''échéance n''est pas renseignée.
    TFMul(Ecran).Pages.ActivePage := TTabSheet(GetControl('PMODIFS', True));
    SetFocusControl('DATECHENEW');
    Exit;
  end;

  if (GetCheckBoxState('CMSRIB') = cbChecked) And (GetControlText('RIBNEW')='') then begin
    HM.Execute(3,'','');  // La nouvelle valeur de RIB n''est pas renseignée.
    TFMul(Ecran).Pages.ActivePage := TTabSheet(GetControl('PMODIFS', True));
    SetFocusControl('RIBNEW');
    Exit;
  end;

  if (FListe.NbSelected > 0) or (Fliste.AllSelected) then begin
    if FListe.AllSelected then begin
      Q.DisableControls;
      Q.First;
      while not Q.Eof do begin
        Transactions(ModifSerieChamp, 2);
        Q.Next;
      end;
      Q.EnableControls;
    end
    else begin
      for i:=0 to Fliste.NbSelected-1 do begin
        FListe.GotoLeBookMark(i);
        {$IFDEF EAGLCLIENT}
        TFMul(Ecran).Q.TQ.Seek(FListe.Row - 1);
        {$ENDIF}
        Transactions(ModifSerieChamp, 2)
      end;
    end;
    TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 13/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CPMODIFECHEMP.AuxiElipsisClick( Sender : TObject );
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;

{b FP 15/06/2006 FQ16972 idem UTOFMULRELCPT}
procedure TOF_CPMODIFECHEMP.FetchLesTous;
begin
  // Applique les critères s'ils ne le sont pas
  if TFMul(Ecran).CritModified then
  begin
    bgFetchAll := True;
    TFMul(Ecran).BChercheClick(Nil);
    Exit;
  end;

  // Récupère tous les enregistrement s'ils ne sont pas tous présents
  if bgFetchAll then
  begin
{$IFDEF EAGLCLIENT}
    TFMul(Ecran).FetchLesTous;
{$ENDIF}
    bgFetchAll := False;
  end;
end;
{e FP 15/06/2006 FQ16972}
Initialization
  registerclasses ( [ TOF_CPMODIFECHEMP ] );
end.



