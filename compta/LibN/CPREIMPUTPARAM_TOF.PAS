{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 26/06/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPREIMPUTPARAM ()
Mots clefs ... : TOF;CPREIMPUTPARAM
*****************************************************************}
Unit CPREIMPUTPARAM_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     UTOB,      // TQuery
{$IFDEF EAGLCLIENT}
     MaineAGL,
{$ELSE}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     FE_Main,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     Ent1,        // VH
     HTB97,       // TToolBarButton97
     SaisUtil,    // ControleDate
     Hcompte,     // TGGeneral, TGTiers
     SaisComm,    // TSAJAL
     LookUp,      // LookupValueExist
     UtilSais,    // NatureGenAuxOK, NaturePieceCompteOk
     AGLInit,     // TheTob
     ULibWindows, // CGereEnabledControl
     {$IFDEF MODENT1}
     CPTypeCons,
     {$ENDIF MODENT1}
     ParamSoc,		// GetParamSocSecur YMO
     UTOF ;

procedure CC_LanceFicheReimpParam;

Type
  TOF_CPREIMPUTPARAM = Class (TOF)
    HM             : THMsgBox;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure InitMsgBox;

    procedure CONTRE_JOURNALExit(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject);
    procedure DirectClick(Sender: TObject);
    procedure AuxiElipsisClick(Sender : TObject);
  private
    MultiDEV : boolean;
    MemeVentil, EncON : Boolean;
    gbDeviseNotPivot : Boolean;
    Gen2Vent   : Array[1..MaxAxe] of boolean ;
    gszAuxiliaire, gszGeneral, gszDevise : String;
    function  JePeuxValider : boolean;
    function  SoucheOK : boolean;
    function  DateOK : boolean;
    function  ComptesOK : boolean;
    procedure Enabled(Value : Boolean);
  end ;

Implementation

uses
    {$IFDEF eAGLCLIENT}
    MenuOLX
    {$ELSE}
    MenuOLG
    {$ENDIF eAGLCLIENT}
    , UTofMulParamGen; {13/04/07 YMO F5 sur Auxiliaire }

procedure CC_LanceFicheReimpParam;
begin
  AGLLanceFiche('CP','CPREIMPUTPARAM','','','');
end;

procedure TOF_CPREIMPUTPARAM.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPREIMPUTPARAM.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPREIMPUTPARAM.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPREIMPUTPARAM.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CPREIMPUTPARAM.OnArgument (S : String ) ;
begin
  Inherited ;
  //SG6 20/01/05 FQ 14910
  Ecran.HelpContext := 7676000;

  // Création des contrôles
  HM := THMsgBox.create(FMenuG) ;
  InitMsgBox;

  //////////////////////////////////////////////////////////////////
  // FormShow
  SetControlText('CONTRE_DATE', DateToStr(V_PGI.DateEntree));
  if Not VH^.MontantNegatif then begin
    SetControlChecked('CONTRE_NEGATIF', False);
    SetControlEnabled('CONTRE_NEGATIF', False);
  end;
  SetControlText('CONTRE_ETABLISSEMENT', VH^.EtablisDefaut);
  MultiDEV := False;
  gbDeviseNotPivot := False;

  // Evénements des contrôles
  TToolBarButton97(GetControl('BVALIDER', True)).OnClick := BOuvrirClick;
  THValComboBox(GetControl('CONTRE_JOURNAL', True)).OnExit := CONTRE_JOURNALExit;
  TCheckBox(GetControl('CHK_DIRECT', True)).OnClick := DirectClick;
  DirectClick(nil);
  SetControlEnabled('CHK_DIRECT', True); // 14472

  // Récupère les paramètres
  gbDeviseNotPivot := (LaTOB.GetValue('GBDEVISENOTPIVOT') = 'X');
  gszAuxiliaire := LaTOB.GetValue('GSZAUXILIAIRE'); // utilisé dans ComptesOK
  gszGeneral := LaTOB.GetValue('GSZGENERAL');       // utilisé dans ComptesOK
  gszDevise := LaTOB.GetValue('GSZDEVISE');         // utilisé dans ComptesOK

  {JP 01/08/05 : FQ 15124 : S'il y a des TTC Gescom, on n'autorise pas la réimputation directe}
  TCheckBox(GetControl('CHK_DIRECT')).Enabled := LaTob.GetString('CANDIRECT') = 'X';
  {Par Précaution ...}
  if LaTob.GetString('CANDIRECT') = '-' then TCheckBox(GetControl('CHK_DIRECT')).Checked := False;

  TheTOB := nil;

  if GetParamSocSecur('SO_CPMULTIERS', false) then
    THEdit(GetControl('CONTRE_AUXILIAIRE', true)).OnElipsisClick:=AuxiElipsisClick;

end ;

procedure TOF_CPREIMPUTPARAM.OnClose ;
begin
  Inherited ;
  HM.Free;
end ;

procedure TOF_CPREIMPUTPARAM.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPREIMPUTPARAM.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_CPREIMPUTPARAM.InitMsgBox;
begin
HM.Mess.Add('0;'+Ecran.Caption+';Vous devez renseigner un journal dans les paramètres.;W;O;O;O;');
HM.Mess.Add('1;'+Ecran.Caption+';La date que vous avez renseignée n''est pas valide.;W;O;O;O;');
HM.Mess.Add('2;'+Ecran.Caption+';La date que vous avez renseignée n''est pas dans un exercice ouvert.;W;O;O;O;');
HM.Mess.Add('3;'+Ecran.Caption+';La date que vous avez renseignée est antérieure à une clôture.;W;O;O;O;');
HM.Mess.Add('4;'+Ecran.Caption+';La date que vous avez renseignée est antérieure à une clôture.;W;O;O;O;');
HM.Mess.Add('5;'+Ecran.Caption+';La date que vous avez renseignée est en dehors des limites autorisées.;W;O;O;O;');
HM.Mess.Add('6;'+Ecran.Caption+';Le compte général que vous avez renseigné est absent ou erroné.;W;O;O;O;');
HM.Mess.Add('7;'+Ecran.Caption+';Le compte auxiliaire que vous avez renseigné est erroné.;W;O;O;O;');
HM.Mess.Add('8;'+Ecran.Caption+';Les natures des comptes que vous avez renseignés sont incompatibles.;W;O;O;O;');
HM.Mess.Add('9;'+Ecran.Caption+';Le compte général que vous avez renseigné est interdit sur le journal.;W;O;O;O;');
HM.Mess.Add('10;'+Ecran.Caption+';Le compte général de banque ne supporte pas la devise de la pièce.;W;O;O;O;');
HM.Mess.Add('11;'+Ecran.Caption+';Le compte général que vous avez renseigné est interdit pour cette nature de pièce.;W;O;O;O;');
HM.Mess.Add('12;'+Ecran.Caption+';Le journal que vous avez renseigné n''est pas multi-devises.;W;O;O;O;');
HM.Mess.Add('13;'+Ecran.Caption+';Le journal que vous avez renseigné ne possède pas de facturier.;W;O;O;O;');
HM.Mess.Add('14;'+Ecran.Caption+';Le couple général/auxiliaire des écritures est incompatible avec celui de contre-passation.;W;O;O;O;');
end;

procedure TOF_CPREIMPUTPARAM.CONTRE_JOURNALExit(Sender: TObject);
Var
  Q : TQuery;
begin
  if GetControlText('CONTRE_JOURNAL')= '' then Exit ;
  MultiDEV:=False ;
  Q:=OpenSQL('SELECT J_MULTIDEVISE FROM JOURNAL WHERE J_JOURNAL="'+GetControlText('CONTRE_JOURNAL')+'"',True);
  if Not Q.EOF then MultiDEV := (Q.Fields[0].AsString='X');
  Ferme(Q);
end;

procedure TOF_CPREIMPUTPARAM.BOuvrirClick(Sender: TObject);
var
  T : TOB;
  i : Integer;
begin
  if Not JePeuxValider then Exit;

  // Renvoi les paramètres
  T := TOB.Create ('', nil, -1);
  T.AddChampSupValeur('DIRECT',GetControlText('CHK_DIRECT'));
  T.AddChampSupValeur('CONTRE_GENERAL',GetControlText('CONTRE_GENERAL'));
  T.AddChampSupValeur('CONTRE_AUXILIAIRE',GetControlText('CONTRE_AUXILIAIRE'));
  T.AddChampSupValeur('CONTRE_JOURNAL',GetControlText('CONTRE_JOURNAL'));
  T.AddChampSupValeur('CONTRE_ETABLISSEMENT',GetControlText('CONTRE_ETABLISSEMENT'));
  T.AddChampSupValeur('CONTRE_DATE',GetControlText('CONTRE_DATE'));
  T.AddChampSupValeur('CONTRE_NEGATIF',GetControlText('CONTRE_NEGATIF'));
  T.AddChampSupValeur('MEMEVENTIL',MemeVentil);
  T.AddChampSupValeur('ENCON',EncON);
  for i := 1 to MaxAxe do begin                            // SBO Correction FFF 127
    T.AddChampSupValeur('GEN2VENT'+IntToStr(i), Gen2Vent[i]);
  end;
  TheTOB := T;
  Ecran.Close;

  // bpy le 14/10/2004
  PGIBox('Il est possible que les ventilations analytiques ne soient plus correctes.','ATTENTION : ');
  // Fin BPY
end;

function TOF_CPREIMPUTPARAM.JePeuxValider: boolean;
begin
  Result := False ;
  if GetCheckBoxState('CHK_DIRECT') = cbUnChecked then begin
    if GetControlText('CONTRE_JOURNAL') = '' then begin HM.Execute(0,'',''); Exit; end; // Vous devez renseigner un journal dans les paramètres.
    if gbDeviseNotPivot and (Not MultiDEV) then begin HM.Execute(12,'',''); Exit; end;  // Le journal de ré-imputation que vous avez renseigné n'est pas multi-devises.
    if Not SoucheOK then Exit;
  end;
  if Not DATEOK then Exit;
  if Not ComptesOK then Exit;
  Result := True;
end;

function TOF_CPREIMPUTPARAM.ComptesOK: boolean;
var
  Gen1,Gen2 : TGGeneral ;
  Aux1,Aux2 : TGTiers ;
  Okok,OkVent,OkLocal : boolean ;
  k,ErrLoc  : integer ;
  SAJAL     : TSAJAL ;
begin
  Result := False;
  Aux1:=Nil; Aux2:=Nil; SAJAL:=Nil;
  MemeVentil := False;
  FillChar(Gen2Vent,Sizeof(Gen2Vent),#0);
  OkVent:=False;

  {Présences obligatoires}
  if GetControlText('CONTRE_GENERAL') = '' then begin
    HM.Execute(6,'','');  // Le compte général de contre-passation que vous avez renseigné est absent ou erroné.
    Exit;
  end;
  if not LookupValueExist(GetControl('CONTRE_GENERAL', True))then begin
    HM.Execute(6,'','');
    Exit;
  end;
  if GetControlText('CONTRE_AUXILIAIRE')<>'' then begin
    if not LookupValueExist(GetControl('CONTRE_AUXILIAIRE')) then begin
      HM.Execute(7,'',''); // Le compte auxiliaire de contre-passation que vous avez renseigné est erroné.
      Exit;
      end
    else
      Aux2 := TGTiers.Create(GetControlText('CONTRE_AUXILIAIRE'));
  end;
  if gszAuxiliaire<>'' then Aux1:=TGTiers.Create(gszAuxiliaire);

  Gen1:=TGGeneral.Create(gszGeneral);
  Gen2:=TGGeneral.Create(GetControlText('CONTRE_GENERAL'));
  if Gen2<>Nil then EncON := Gen2.TvaSurEncaissement;

  if GetControlText('CONTRE_JOURNAL')<>'' then SAJAL:=TSAJAL.Create(GetControlText('CONTRE_JOURNAL'),False);
  Okok := True;

  {Cohérence comptes de sortie}
  OkLocal := True;
  ErrLoc := 0;

  if ((Gen2<>Nil) and (Aux2<>Nil)) then begin
    if Not NatureGenAuxOk(Gen2.NatureGene,Aux2.NatureAux) then begin
      ErrLoc := 8;
      OkLocal := False;
    end;
  end;

  if ((SAJAL<>Nil) and (Gen2<>Nil)) then
    if EstInterdit(SAJAL.COMPTEINTERDIT,Gen2.General,0)>0 then begin
      ErrLoc := 9;
      OkLocal := False;
    end;

  if (Gen2<>Nil) then begin
    if Gen2.NatureGene='BQE' then
      if PasDeviseBanque(Gen2.General,gszDevise) then begin
        ErrLoc := 10;
        OkLocal := False;
      end;
  end;

  if Gen2<>Nil then
    if Not NaturePieceCompteOk('OD',Gen2.NatureGene) then begin
      ErrLoc := 11;
      OkLocal := False;
    end;

  if Not OkLocal then begin
    Gen1.Free; Gen2.Free; Aux1.Free; Aux2.Free; SAJAL.Free;
    HM.Execute(ErrLoc,'','');
    Exit;
  end;

  {Cohérence 1 et 2}
  if ((Gen1=Nil) or (Gen2=Nil)) then Okok := False
  else if ((Aux1<>Nil) and (Aux2=Nil)) or ((Aux1=Nil) and (Aux2<>Nil)) then Okok := False
  else if Gen1.Lettrable<>Gen2.Lettrable then Okok := False
  else if Gen1.Collectif<>Gen2.Collectif then Okok := False
  else if ((Gen1.Collectif) and (Aux1=Nil)) or ((Gen2.Collectif) and (Aux2=Nil)) then Okok := False
  else if ((Not Gen1.Collectif) and (Aux1<>Nil)) or ((Not Gen2.Collectif) and (Aux2<>Nil)) then Okok := False
  else begin
    if ((Aux1<>Nil) and (Aux2<>Nil)) then begin
      if Aux1.Lettrable<>Aux2.Lettrable then Okok := False
      else if ((Gen1.General=Gen2.General) and (Aux1.Auxi=Aux2.Auxi)) then Okok := False;
    end;

    if Okok then begin
      if Gen1.General=Gen2.General then begin
        if ((Aux1=Nil) and (Aux2=Nil)) or ((Aux1.Auxi=Aux2.Auxi)) then Okok := False;
      end;
      MemeVentil := True;

      for k:=1 to MaxAxe do begin
        if Gen1.Ventilable[k]<>Gen2.Ventilable[k] then MemeVentil:=False;
        Gen2Vent[k] := Gen2.Ventilable[k];
        if Gen1.Ventilable[k] then OkVent := True;
      end;
      if Not OKVent then MemeVentil := False;
    end;
  end;

  {Fin tests}
  Result := Okok;
  Gen1.Free; Gen2.Free; Aux1.Free; Aux2.Free; SAJAL.Free;
  if Not Okok then HM.Execute(14,'','');  // Le couple général / auxiliaire des écritures est incompatible avec celui de contre-passation.
end;

function TOF_CPREIMPUTPARAM.DateOK: boolean;
var
  Err : integer ;
begin
  Result := False;
  Err := ControleDate(GetControlText('CONTRE_DATE'));
  if Err>0 then begin
    HM.Execute(Err,'','');  // Erreur 1 à 5
    Exit;
  end;
  Result := True;
end;

function TOF_CPREIMPUTPARAM.SoucheOK: boolean;
var
  Q : TQuery ;
begin
  Result := False;
  if GetControlText('CONTRE_JOURNAL') = '' then Exit;
  Q := OpenSQL('SELECT J_COMPTEURNORMAL FROM JOURNAL WHERE J_JOURNAL="'+GetControlText('CONTRE_JOURNAL')+'"',True);
  if Not Q.EOF then begin
    if Q.Fields[0].AsString<>'' then Result := True;
  end;
  Ferme(Q);
  if Not Result then HM.Execute(13,'','');  // Le journal que vous avez renseigné ne possède pas de facturier.
end;

procedure TOF_CPREIMPUTPARAM.DirectClick(Sender: TObject);

begin
  if GetCheckBoxState('CHK_DIRECT') = cbChecked then Enabled(False)
                                                else Enabled(True);
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 13/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CPREIMPUTPARAM.AuxiElipsisClick( Sender : TObject );
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;


procedure TOF_CPREIMPUTPARAM.Enabled(Value: Boolean);
var
  i : Integer;
  C : TCustomControl;
begin
  C := TCustomControl(GetControl('GBX_NODIRECT', True));
  if C = Nil then Exit;

  for i:=0 to C.ControlCount - 1 do begin
    if C.Controls[i].Tag <> 1 then C.Controls[i].Enabled := Value;
  end;
end;

Initialization
  registerclasses ( [ TOF_CPREIMPUTPARAM ] ) ;
end.
