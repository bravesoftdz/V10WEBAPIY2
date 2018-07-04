{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 08/07/2002
Modifié le ... : 08/07/2002
Description .. : Source TOF de la FICHE : LGCOMPTE ()
Mots clefs ... : TOF;LGCOMPTE
*****************************************************************}
Unit LGCOMPTE_TOF ;

Interface

Uses Windows, StdCtrls, Controls, Classes,
{$IFDEF EAGLCLIENT}
     MAinEAGL,
{$ELSE}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     FE_Main,
{$ENDIF}
     forms, sysutils,
     HCtrls, HEnt1, Ent1, UTob, 
     HMsgBox,
     UTOF,
     Paramsoc,
     Spin,
     UtilPGI,
     Dialogs,
     CLgCpte ;

procedure CCLanceFiche_LongueurCompte;

Type
  TOF_LGCOMPTE = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnLoad                   ; override ;
    procedure OnClose                  ; override ;
		procedure OnUpdate                 ; override ;
  private
    FBourre : Char;
    FLg     : integer;
    FLCpt   : TStringList;
    gszAxe  : String;
    procedure InitZone (St : string);
    procedure ChangeLongueurGen;
    procedure ChangeLongueurAux;
    procedure ChangeLongueurAna;
    procedure OnInfoMessage ( Msg : string);
    procedure AfficheMessageInfo (Entete,Detail : string);
    function  JaiLeDroitChangerLongueur : boolean;
    procedure OnFormKeyDown (Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure SupprimeLesComptes ( Table : string );
    procedure cbo_AxeChange(Sender: TObject);
  end ;

Const	MsgBox: array[0..22] of string = (
	{0} 'Confirmez-vous le changement de longueur des comptes généraux ?',
	{1}	'Recherche des éventuels doublons en cours ...',
	{2}	'Traitement impossible.#10#13Cette opération va générer des doublons dans la table des comptes généraux.#10#13Incompatibilité de paramétrage sur le compte ',
	{3}	'Mise à jour des comptes généraux en cours ...',
	{4}	'Changement de longueur des comptes généraux impossible.',
	{5}	'Rechargement des paramètres société en cours ...',
	{6}	'Changement de longueur des comptes généraux terminé avec succès.',
	{7} 'Confirmez-vous le changement de longueur des comptes auxiliaires ?',
  {8} 'Traitement impossible.#10#13Cette opération va générer des doublons dans la table des comptes auxiliaires.#10#13Incompatibilité de paramétrage sur le compte ',
  {9} 'Mise à jour des comptes auxiliaires en cours ...',
 {10} 'Changement de longueur des comptes auxiliaires impossible.',
 {11} 'Changement de longueur des comptes auxiliaires terminé avec succès.',
 {12} 'Confirmez-vous le changement de longueur des sections analytiques ?',
 {13} 'Traitement impossible.#10#13Cette opération va générer des doublons dans la table des sections analytiques.#10#13Incompatibilité de paramétrage sur la section ',
 {14} 'Mise à jour des sections analytiques en cours ...',
 {15} 'Changement de longueur des sections analytiques impossible.',
 {16} 'Changement de longueur des sections analytiques terminé avec succès.',
 {17} 'Traitement impossible.#10#13L''axe est structuré.',
 {18} 'Cette fonctionnalité n''est pas disponible sur Interbase.',
 {19} 'Ce traitement va modifier les paramètres de base de votre société.#10#13Assurez-vous que vous avez une sauvegarde de la base de données avant de lancer le traitement.',
 {20} 'Vous ne pouvez pas diminuer la longueur des comptes.#10#13Traitement impossible, dossier est en mode synchonisation.',
 {21} 'Cette fonctionnalité n''est pas disponible en multi société.',
 {22} 'Vous ne pouvez pas changer le paramètrage de l''axe de TVA.');

Implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  {$ENDIF MODENT1}
  uProcGen;

procedure CCLanceFiche_LongueurCompte;
begin
  if V_PGI.Driver=dbINTRBASE then PGIInfo(MsgBox[18],'Changement des longueurs des comptes')
  else if EstMultiSoc then PGIInfo (MsgBox[21],'Changement des longueurs des comptes')
  else AGLLanceFiche('CP','CPLGCOMPTE','','','') ;
end;

procedure TOF_LGCOMPTE.InitZone(St: string);
begin
  if GetControl(St) is TCheckBox then
  begin
    SetControlChecked(St,GetParamSocSecur(St,False)) ;
  end
  else
  begin
      if GetControl(St) is TSpinEdit then
        SetControlText(St,GetParamSocSecur(St,0))
      else
        SetControlText(St,GetParamSocSecur(St,''))
  end;
end;

procedure TOF_LGCOMPTE.OnArgument (S : String ) ;
var
	C : THValComboBox;
begin
  Inherited ;

  Ecran.HelpContext:=999999443;

  C := THValComboBox(GetControl('CBO_AXE'));	if (C <> nil) then C.OnChange := cbo_AxeChange;
  AfficheMessageInfo('','');
  SetControlProperty('SO_LGCPTEGEN','MaxValue',17);
  SetControlProperty('SO_LGCPTEGEN','MinValue',6);
  SetControlProperty('SO_LGCPTEAUX','MaxValue',17);
  SetControlProperty('SO_LGCPTEAUX','MinValue',6);
  SetControlProperty('CBO_AXE','Value','A1');
  SetControlProperty('X_LONGSECTION','MaxValue',17);
  SetControlProperty('X_LONGSECTION','MinValue',3);
  if ctxPCL in V_PGI.PGIContexte then SetControlEnabled('X_BOURREANA',False);
  Ecran.OnKeyDown := OnFormKeyDown;
  FLCpt := TStringList.Create;

  // Groupe box de l'analytique
  SetControlVisible('FE__GROUPBOX', True);
end ;

procedure TOF_LGCOMPTE.OnLoad ;
begin
  Inherited ;
  InitZone('SO_LGCPTEGEN');
  InitZone('SO_BOURREGEN');
  InitZone('SO_LGCPTEAUX');
  InitZone('SO_BOURREAUX');

  // GCO - 05/09/2006 - FQ 18741 - la ligne était mise en commentaire sans
  // explication par l'auteur de la modification, je la réactive donc.
  cbo_AXEChange(Nil);
end ;

procedure TOF_LGCOMPTE.OnClose ;
begin
  Inherited ;
  DeBlocageMonoPoste(True);
  FLCpt.Free;
end ;

procedure TOF_LGCOMPTE.ChangeLongueurAux;
begin
  // Change la longueur des comptes auxiliaires
  SupprimeLesComptes ( 'TIERS' );
  ChangeLgCpteAux ( FLg , FBourre , OnInfoMessage );
end;

procedure TOF_LGCOMPTE.ChangeLongueurGen;
begin
  // Change la longueur des comptes généraux
  SupprimeLesComptes ( 'GENERAUX' );
  ChangeLgCpteGen ( FLg , FBourre , OnInfoMessage);
end;

procedure TOF_LGCOMPTE.ChangeLongueurAna;
begin
  // Change la longueur des sections analytiques
  SupprimeLesComptes ( 'SECTION' );
  ChangeLgCpteAna ( FLg , FBourre , gszAxe, OnInfoMessage );
end;

procedure TOF_LGCOMPTE.OnUpdate ;
var Bourre : Char;
    Lg : integer;
    OkMaj : boolean;
begin
  if Not JaiLeDroitChangerLongueur then exit;
  if Not BlocageMonoPoste(True) then Exit ;

  // Mode synchronisation
  if GetParamsocSecur('SO_CPMODESYNCHRO', FALSE) and LienS1 then
  begin
    if (((TSpinEdit(GetControl('SO_LGCPTEGEN'))).value < GetParamSocSecur('SO_LGCPTEGEN',0)) or
    ((TSpinEdit(GetControl('SO_LGCPTEAUX'))).value < GetParamSocSecur('SO_LGCPTEAUX',0)) or
    ((TSpinEdit(GetControl('X_LONGSECTION'))).value < VH^.Cpta[AxeToFb(gszAxe)].Lg)) then
    begin
       (TSpinEdit(GetControl('SO_LGCPTEGEN'))).value := GetParamSocSecur('SO_LGCPTEGEN',0);
       (TSpinEdit(GetControl('SO_LGCPTEAUX'))).value := (GetParamSocSecur('SO_LGCPTEAUX',0));
       (TSpinEdit(GetControl('X_LONGSECTION'))).value := VH^.Cpta[AxeToFb(gszAxe)].Lg;
       PGIInfo (MsgBox[20], Ecran.Caption);
       exit;
    end;
  end;

  if (not (ctxPCL in V_PGI.PGIContexte)) and
    (((TSpinEdit(GetControl('SO_LGCPTEGEN'))).value <> GetParamSocSecur('SO_LGCPTEGEN',0)) or
    ((TSpinEdit(GetControl('SO_LGCPTEAUX'))).value <> GetParamSocSecur('SO_LGCPTEAUX',0)) or
    ((TSpinEdit(GetControl('X_LONGSECTION'))).value<>VH^.Cpta[AxeToFb(gszAxe)].Lg)) then
    PGIInfo(MsgBox[19],Ecran.Caption);

  // Changement de longueur des comptes généraux
  Bourre := StrToChr(GetControlText('SO_BOURREGEN')[1]);
  Lg := (TSpinEdit(GetControl('SO_LGCPTEGEN'))).value;
  if Lg <> GetParamSocSecur('SO_LGCPTEGEN',0) then
  begin
    if PGIAsk(MsgBox[0],Ecran.Caption)<>mrYes then
    begin
      DeblocageMonoPoste(True) ;
      exit;
    end else
    begin
      AfficheMessageInfo (MsgBox[1],'');
      OkMaj := PasDeDoublonChampCompte ('GENERAUX', 'G_GENERAL',Bourre, Lg, FLCpt);
      if not OkMaj then
      begin
        AfficheMessageInfo ('','');
        PGIBox(MsgBox[2]+FLCpt.Strings[FLCpt.Count-1]+'.',Ecran.Caption);
        FLCpt.Clear;
        DeblocageMonoPoste(True) ;
        exit;
      end else
      begin
        FBourre := Bourre;
        FLg := Lg;
        AfficheMessageInfo (MsgBox[3],'');
        if (Transactions(ChangeLongueurGen,1) <>oeOK) then
        begin
          MessageAlerte(MsgBox[4]);
          AfficheMessageInfo ('','');
          DeblocageMonoPoste(True) ;
          exit;
        end else
        begin
          SetParamSoc('SO_LGCPTEGEN',Lg);
          SetParamSoc('SO_BOURREGEN',Bourre);
          AfficheMessageInfo (MsgBox[5],'');
        // GCO - 25/09/2007 - FQ 21277
        {$IFDEF EAGLCLIENT}
          AvertirCacheServer('PARAMSOC');
        {$ENDIF}
          RechargeParamSoc;
          AfficheMessageInfo ('','');
          PGIInfo(MsgBox[6],Ecran.Caption);
        end;
      end;
    end;
  end;

  // Changement de longueur des comptes auxiliaires
  Bourre := StrToChr(GetControlText('SO_BOURREAUX')[1]);
  Lg := (TSpinEdit(GetControl('SO_LGCPTEAUX'))).value;
  if Lg <> GetParamSocSecur('SO_LGCPTEAUX',0) then
  begin
    if PGIAsk(MsgBox[7],Ecran.Caption)<>mrYes then
    begin
      DeblocageMonoPoste(True) ;
      exit;
    end else
    begin
      AfficheMessageInfo (MsgBox[1],'');
      OkMaj := PasDeDoublonChampCompte ('TIERS', 'T_AUXILIAIRE',Bourre, Lg, FLCpt);
      if not OkMaj then
      begin
        AfficheMessageInfo ('','');
        PGIBox (MsgBox[8]+FLCpt.Strings[FLCpt.Count-1]+'.',Ecran.Caption);
        FLCpt.Clear;
        DeblocageMonoPoste(True) ;
        exit;
      end else
      begin
        FBourre := Bourre;
        FLg := Lg;
        AfficheMessageInfo (MsgBox[9],'');
        if (Transactions(ChangeLongueurAux,1) <>oeOK) then
        begin
          MessageAlerte(MsgBox[10]);
          AfficheMessageInfo ('','');
          DeblocageMonoPoste(True) ;
          exit;
        end else
        begin
          SetParamSoc('SO_LGCPTEAUX',Lg);
          SetParamSoc('SO_BOURREAUX',Bourre);
          AfficheMessageInfo (MsgBox[5],'');
        // GCO - 25/09/2007 - FQ 21277
        {$IFDEF EAGLCLIENT}
          AvertirCacheServer('PARAMSOC');
        {$ENDIF}
          RechargeParamSoc;
          AfficheMessageInfo ('','');
          PGIInfo(MsgBox[11],Ecran.Caption);
        end;
      end;
    end;
  end;

  // Changement de longueur des sections analytiques
  Bourre := StrToChr(GetControlText('X_BOURREANA')[1]);
  Lg := (TSpinEdit(GetControl('X_LONGSECTION'))).value;

  if Lg <> VH^.Cpta[AxeToFb(gszAxe)].Lg then begin
    if PGIAsk(MsgBox[12],Ecran.Caption)<>mrYes then
    begin
      exit;
    end else
    begin
      AfficheMessageInfo (MsgBox[1],'');
      if GetColonneSQL('AXE','X_STRUCTURE','X_AXE="'+gszAxe+'"')='X' then
      begin
        AfficheMessageInfo ('','');
        PGIBox (MsgBox[17],Ecran.Caption);
        DeblocageMonoPoste(True) ;
        exit;
      end else
      if GetParamSocSecur('SO_CPPCLSAISIETVA',false) and ( THValComboBox(GetControl('CBO_AXE')).Value = GetParamSocSecur('SO_CPPCLAXETVA', '') ) then
      begin
        AfficheMessageInfo ('','');
        PGIBox (MsgBox[22],Ecran.Caption);
        DeblocageMonoPoste(True) ;
        exit;
      end ;
      OkMaj := PasDeDoublonChampCompte ('SECTION', 'S_SECTION',Bourre, Lg, FLCpt, gszAxe);
      if not OkMaj then
      begin
        AfficheMessageInfo ('','');
        PGIBox (MsgBox[13]+FLCpt.Strings[FLCpt.Count-1]+'.',Ecran.Caption);
        FLCpt.Clear;
        DeblocageMonoPoste(True) ;
        exit;
      end else
      begin
        FBourre := Bourre;
        FLg := Lg;
        AfficheMessageInfo (MsgBox[14],'');
        if (Transactions(ChangeLongueurAna,1) <>oeOK) then
        begin
          MessageAlerte(MsgBox[15]);
          AfficheMessageInfo ('','');
          DeblocageMonoPoste(True) ;
          exit;
        end else
        begin
          ExecuteSQL('UPDATE AXE SET X_LONGSECTION='+IntToStr(LG)+',X_BOURREANA="'+Bourre+'" WHERE X_AXE="'+gszAxe+'"');
          AfficheMessageInfo (MsgBox[5],'');
          CHARGEMAGHALLEY;
          AfficheMessageInfo ('','');
          PGIInfo(MsgBox[16],Ecran.Caption);
        end;
      end;
    end;
  end;

  AfficheMessageInfo ('','');
  inherited;
end ;

procedure TOF_LGCOMPTE.OnInfoMessage(Msg: string);
begin
  SetControlProperty('INFOMESSAGEDETAIL','Caption',Msg);
  Application.ProcessMessages;
end;

procedure TOF_LGCOMPTE.AfficheMessageInfo(Entete, Detail: string);
begin
  SetControlProperty('INFOMESSAGE','Caption',Entete);
  SetControlProperty('INFOMESSAGEDETAIL','Caption',Detail);
  Application.ProcessMessages;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/09/2006
Modifié le ... :   /  /    
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_LGCOMPTE.JaiLeDroitChangerLongueur: boolean;
var {St , }Msg : string;
{    Q : TQuery;
    bOk : boolean;}
begin

  if not (ctxPCL in V_PGI.PGIContexte) then begin Result := True; Exit; end;
  Result := False;
  if (not IsMSSQL) then  //10/08/2006 YMO Ajout test SQL2005
    Msg := 'Manipulation autorisée uniquement avec SQL Server.'
  else if not (ctxPCL in V_PGI.PGIContexte) then
    Msg := 'Manipulation autorisée uniquement dans un contexte PCL.'
  else if ((ctxPCL in V_PGI.PGIContexte) and (EstBasePCLAllegee and (not EstBasePclOptimisee))) then
    Msg := 'Vous devez optimiser votre dossier pour lancer ce traitement.'
  else if ((not EstBasePclAllegee) and  ExisteSQL('SELECT 1 FROM COMPTADIFFEREE WHERE GCD_REFPIECE <>""')) then
    Msg := 'Traitement impossible : présence de pièces commerciales à envoyer en comptabilité.'
{  else if V_PGI_Env.InBaseCommune then
    Msg := 'Manipulation non autorisée dans la base commune.'}
  else Result := True;
{
  else
  begin
    bOk := True;
    Q := TQuery.Create(nil);
    Q.DatabaseName := 'DBCOM';
    Q.SQL.Text := 'SELECT DAP_NODOSSIER,DAP_NOMEXEC FROM DOSSAPPLI WHERE DAP_NODOSSIER="' +
      V_PGI_Env.NoDossier + '"';
    ChangeSQL(Q);
    Q.RequestLive := True;
    Q.Open;
    while not Q.Eof do
    begin
      St := Q.FindField('DAP_NOMEXEC').AsString;
//      if ((St <> 'CCS5.EXE') and (St<>'CPS5.EXE')) then bOk := False;
// CA - 17/06/2002 - Changement de longueur des comptes interdit
// uniquement si GI activée
      if (St = 'CGIS5.EXE')  then bOk := False;
      Q.Next;
    end;
    Q.Close;
    Q.Free;
    if not bOk then
      Msg := 'Manipulation non autorisée : la Gestion Interne a été activée sur ce dossier.'
    else Result := True;
  end;}
  if not Result then MessageAlerte (Msg);
end;

procedure TOF_LGCOMPTE.OnFormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
 if ( csDestroying in Ecran.ComponentState ) then Exit ;
  case Key of
    VK_F10 :
      begin
        TButton(GetControl('BVALIDER')).Click;
        Key := 0;
      end;
  end;
end;

procedure TOF_LGCOMPTE.SupprimeLesComptes(Table: string);
var i : integer;
    Champ : string;
begin
	if Table = 'SECTION' then begin
	  for i:=0 to FLCpt.Count - 1 do
  	  ExecuteSQL ('DELETE FROM '+Table+' WHERE '+Champ+'="'+FLCpt.Strings[i]+'" AND S_AXE = "'+gszAxe+'"');
	  FLCpt.Clear;
    exit
  end;

  if Table = 'TIERS' then Champ := 'T_AUXILIAIRE'
  else Champ := 'G_GENERAL';
  for i:=0 to FLCpt.Count - 1 do
    ExecuteSQL ('DELETE FROM '+Table+' WHERE '+Champ+'="'+FLCpt.Strings[i]+'"');
  FLCpt.Clear;
end;

procedure TOF_LGCOMPTE.cbo_AxeChange(Sender: TObject);
var
 Q      : TQuery ;
 lBoTVA : boolean ;
begin
 gszAxe := THValComboBox(GetControl('CBO_AXE')).Value;

 Q := OpenSQL('SELECT X_LONGSECTION, X_BOURREANA FROM AXE WHERE X_AXE ="'+gszAxe+'"',True,-1, '', True);
  {JP 21/08/07 : FQ 21249 : Ne pas Tester if Q.BOF mais not EOF, car en cwas, sur une TOB ....}
  if not Q.EOF then
   begin
    SetControlProperty('X_LONGSECTION','Value',Q.Fields[0].AsString);
    SetControlText('X_BOURREANA',Q.Fields[1].AsString);
   end;
 Ferme(Q);


 lBoTVA := GetParamSocSecur('SO_CPPCLSAISIETVA',false) and ( THValComboBox(GetControl('CBO_AXE')).Value = GetParamSocSecur('SO_CPPCLAXETVA', '') ) ;

 SetControlProperty('X_LONGSECTION' ,'Enabled' , not lBoTVA) ;
 SetControlProperty('FE_HLABEL'     ,'Enabled' , not lBoTVA) ;
 SetControlProperty('X_BOURRREANA'  ,'Enabled' , not lBoTVA) ;
 SetControlProperty('FE_HLABEL3'    ,'Enabled' , not lBoTVA) ;


end;

Initialization
  registerclasses ( [ TOF_LGCOMPTE ] ) ;
end.
