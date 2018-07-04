{ Unité : Source TOF de la FICHE : CPMULRUB ()
--------------------------------------------------------------------------------------
    Version    |   Date   | Qui  |   Commentaires
--------------------------------------------------------------------------------------
4.3.1.101.001    18/06/03    JP     Création de l'unité
                 09/09/03    JP     Gestion des rubriques pour la Tréso cf IFDEF TRESO
                 13/09/05    GCO    FQ 16629
6.53.001.001     17/01/06    JP     FQ 17304 : En mode comptable classique
6.53.001.001     17/01/06    JP     FQ 17306 : Gestion des listes et de filtres (SetDBListe)
8.00.000.000     23/11/06    JP     FQ 17304 : suite ajout des classes CDR et RDR
                 20/11/07    LM     sur la base de l'argument "V2", on peut définir en fin d'argument une série de critères : AGLLanceFiche('CP', 'CPMULRUB', '', '', 'ACTION=CONSULTATION;V2;NON;RB_TYPERUB=RBF;RB_RUBRIQUE=@TPRO;');
--------------------------------------------------------------------------------------}
Unit CPMULRUB_TOF ;

Interface

uses
     StdCtrls,
     Controls,
     Classes,
    {$IFDEF EAGLCLIENT}
     MaineAGL,
     UTob,
     eMul,
    {$ELSE}
     FE_Main,
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     Mul,
    {$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     {$IFNDEF TRESO}
     CpteUtil,
     {$ENDIF}
     Ent1,
     Hqry,
     AGLInit,
     HTB97,
     RUBRIQUE_TOM, // TTypeFCtxSurRubrique
     uTOF ;

Type
  TOF_CPMULRUB = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure OnDisplay                ; override ;
  private
    Budget,
    RubDeRub : Boolean ;
    FCtxSurRubrique : TCtxSurRubrique;

    RB_SIGNERUB  : THValComboBox;
    RB_TYPERUB   : THValComboBox;
    TRB_SIGNERUB : THLabel;
    LMsg         : THMsgBox;
    TypeAction : TActionFiche; // VL 02/02/2005
    procedure ChargeCbTypRub ;
    procedure RecupListeRub ;
    procedure EnvoiListeRub ;
    procedure FListeDblClick(Sender : TObject);
    procedure FListeOnRowEnter ( Sender : TObject);
    procedure BInsertClick  (Sender : TObject);
    procedure BInsertPersoClick  (Sender : TObject);
    procedure InitMsg ;
    procedure FermeMsg ;
  end ;

procedure MultiCritereRubriqueV2(Comment : TActionFiche ; Budget : Boolean) ;
procedure MultiCritereRubriqueV3(TypRub,Ax,NatRub,BudJal : string ; var ListeRubrique : string ; Budget : Boolean) ;


implementation

uses
    {$IFDEF MODENT1}
    CPProcGen,
    CPTypeCons,
    CPVersion,
    {$ENDIF MODENT1}
    {$IFDEF eAGLCLIENT}
    MenuOLX
    {$ELSE}
    MenuOLG
    {$ENDIF eAGLCLIENT}
    , UtilPGI
    , usatUtil
    ;

var
  ListeRub : string ;

{---------------------------------------------------------------------------------------}
procedure MultiCritereRubriqueV2(Comment : TActionFiche ; Budget : Boolean) ;
{---------------------------------------------------------------------------------------}
var
  ch : string;
begin
  ch := ActionToString(Comment) + ';' + 'V2' + ';';
  if Budget then ch := ch + 'OUI;'
            else ch := ch + 'NON;';
  AGLLanceFiche('CP', 'CPMULRUB', '', '', ch);
end;

{---------------------------------------------------------------------------------------}
Procedure MultiCritereRubriqueV3(TypRub,Ax,NatRub,BudJal : string ; var ListeRubrique : string ; Budget : Boolean) ;
{---------------------------------------------------------------------------------------}
var
  ch : string;
begin
  ch := 'ACTION=CONSULTATION;' + 'V3' + ';' + TypRub + ';' + Ax + ';' + NatRub + ';' + BudJal + ';' +ListeRub + ';';
  if Budget then ch := ch + 'OUI;'
            else ch := ch + 'NON;';
  AGLLanceFiche('CP', 'CPMULRUB', '', '', ch);
  ListeRubrique := ListeRub ;
end ;

////////////////////////////////////////////////////////////////////////////////

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULRUB.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
var
  Appel,
  Filtre,
  TypRub,
  Axe,
  NatRub,
  BudJal : string;
  Fiche  : TFMul;
  st, w, cle, vaut : string ;//LM20071120

begin
  inherited;

  //28/10/2005 Rajout de l'aide
  Ecran.HelpContext:=600200;

  TypeAction := StrToTA(ReadTokenSt(S)); // VL 02/02/2005
  Appel := ReadTokenSt(S);
  InitMsg;
  RubDeRub := False ;
  Budget := False;
  Fiche := TFMul(Ecran);
  Filtre := '';

  if Appel = 'V3' then
  begin
    RubDeRub := True;
    TypRub   := ReadTokenSt(S);
    Axe      := ReadTokenSt(S);
    NatRub   := ReadTokenSt(S);
    BudJal   := ReadTokenSt(S);
    ListeRub := ReadTokenSt(S);
    Budget   := ReadTokenSt(S) = 'OUI';
    if Budget then
    begin
      FCtxSurRubrique := CtxBudget;
      Fiche.FNomFiltre := 'MULVRUBBU';
      Fiche.HelpContext := 15415000 ;
      Filtre :=  'RB_NATRUB="BUD"';
    end
    else
    begin
      FCtxSurRubrique := CtxRubrique;
      Fiche.FNomFiltre := 'MULVRUB';
      Fiche.HelpContext := 7781000 ;
      {JP 17/01/06 : FQ 17304 : Ajout de "TLI"
       JP 23/11/06 : FQ 17304 : fiche réouverte pour l'ajout de CDR et RDR}
      Filtre := 'RB_NATRUB="CPT" AND RB_CLASSERUB IN ("CPT", "TLI", "CDR", "RDR")';
    end;

    if TypRub <> '' then
      Filtre := Filtre + ' AND RB_TYPERUB="' + TypRub + '"';
    if Axe <> '' then
      Filtre := Filtre + ' AND RB_AXE="' + Axe + '"';
    if NatRub <> '' then
      Filtre := Filtre + ' AND RB_NATRUB="' + NatRub + '"';
    if BudJal <> '' then
      Filtre := Filtre + ' AND RB_BUDJAL="' + BudJal + '"';

    Fiche.Caption:= LMsg.Mess[0] ;

//    Fiche.FListe.MultiSelection := True;

    {???????}
    Fiche.BOuvrir.OnClick := nil ;
    Fiche.BOuvrir.ModalResult := mrOK ;
    {???????}
  end
  else
    if Appel = 'V2' then
    begin
      Budget   := ReadTokenSt(S) = 'OUI';
      {$IFDEF TRESO}
      Fiche.FNomFiltre := 'RUBTRESO';
      //RB_TYPERUB : TR pour Tréso, P (Positif), N (Négatif)
      Filtre := 'RB_NATRUB = "TRE" AND (RB_TYPERUB = "TRP" OR RB_TYPERUB = "TRN")';
      {$ELSE}

      if Budget then
      begin
        FCtxSurRubrique := CtxBudget;
        {FQ 17306 : devenu inutile avec la nouvelle gestion des listes
         Fiche.FNomFiltre := 'MULVRUBBU' ;}
        Filtre := 'RB_NATRUB="BUD" ' ;
      end
      else
      begin
        FCtxSurRubrique := CtxRubrique;
        {FQ 17306 : devenu inutile avec la nouvelle gestion des listes
        Fiche.FNomFiltre := 'MULVRUB' ;}
        {JP 17/01/06 : FQ 17304 : Ajout de "TLI"
         JP 23/11/06 : FQ 17304 : fiche réouverte pour l'ajout de CDR et RDR}
        Filtre := 'RB_NATRUB="CPT" AND RB_CLASSERUB IN ("CPT", "TLI", "CDR", "RDR")';
      end;
      {$ENDIF}

      case TypeAction of // VL 02/02/2005
        taConsult : begin
                      Fiche.Caption:= LMsg.Mess[0] ;
                      if FCtxSurRubrique = CtxBudget then Fiche.HelpContext := 15415000
                      else Fiche.HelpContext := 7781000 ;
                    end ;
        taModif   : begin
                      Fiche.Caption:= LMsg.Mess[1] ;
                      if FCtxSurRubrique = CtxBudget then Fiche.HelpContext := 15419000
                      else Fiche.HelpContext := 7787000 ;
                    end ;
      end ; {case}

      //+LM20071120
      if s<>'' then
      begin
        st:=s ;
        repeat
          w:=readTokenst(st) ;
          cle:=trim(gtfs(w,'=',1)) ;
          vaut:=trim(gtfs(w,'=',2)) ;
          if (cle<>'') and (vaut<>'') and (cle<>vaut) and (Ecran.FindComponent(cle)<>nil) then
            setControlText(cle, vaut) ;
        until st=''
      end ;
      //-LM20071120

    end;

(* CA - 18/05/2005 - C'est plutôt le contraire, non ? *)
(*
  if VH^.PaysLocalisation<>CodeISOES then //XVI 24/02/2005
     Filtre:=Filtre+' and (not RB_RUBRIQUE like "@%")' ;
     *)
  if VH^.PaysLocalisation=CodeISOES then //CA 18/05/2005
     Filtre:=Filtre+' and (not RB_RUBRIQUE like "@%")' ;

  Fiche.BInsert.Visible := True ;
  SetControlVisible('BINSERTPERSO',(FCtxSurRubrique=ctxRubrique));
  Fiche.FListe.OnDblClick := FListeDblClick;
  //  SBO 17/01/06 : En attendant finalisation de la fiche, je mets en commentaire pour pouvoir compiler en CWAS
//  Fiche.FListe.OnRowEnter := FListeOnRowEnter;

  {JP 17/01/06 : FQ 17306 : Correction de la gestion des filtres et des listes :
                 création d'une liste pour les budgets}
  {$IFDEF TRESO}
  Fiche.SetDBListe('TRRUBRIQUETRESO');
  {$ELSE}
  if Budget then
    Fiche.SetDBListe('CPRUBRIQUEBUD')
  else
    Fiche.SetDBListe('RUBRIQUE');
  {$ENDIF}

  SetControlText('XX_WHERE', Filtre);

  RB_SIGNERUB  := THValComboBox(GetControl('RB_SIGNERUB'));
  RB_TYPERUB   := THValComboBox(GetControl('RB_TYPERUB'));
  TRB_SIGNERUB := THLabel      (GetControl('TRB_SIGNERUB'));

  TToolbarButton97(GetControl('BINSERT')).OnClick := BInsertClick;
  if Assigned(GetControl('BINSERTPERSO')) then
    TToolbarButton97(GetControl('BINSERTPERSO')).OnClick := BInsertPersoClick;

  Fiche.Caption := LMsg.Mess[8] ;

  {$IFDEF TRESO}
  {Pour la Tréso la gestion des rubriques est simplifiée : elle permet juste de rattacher
   des comptes généraux à des rubriques : les rubriques sont des simili Flux et permettent
   de regrouper les les écritures et de donner un libellé aux groupe de comptes}
  SetControlVisible('RB_SIGNERUB' , False);
  SetControlVisible('TRB_SIGNERUB', False);
  SetControlVisible('RB_TYPERUB'  , False);
  SetControlVisible('TRB_TYPERUB' , False);
  {$ENDIF}
  {JP 09/09/2003 : Remarque pertinente de Laroche, c'est mieux que dans le OnLoad}
  ChargeCbTypRub ;

  Application.ProcessMessages;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULRUB.FListeDblClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  bStdCegid : Boolean;
  Q         : THQuery;
  TmpAction : TActionFiche;
begin
  {$IFDEF EAGLCLIENT}
  if TFMul(Ecran).FListe.RowCount = 0 then Exit;
  {$ELSE}
  if TFMul(Ecran).Q.Eof and TFMul(Ecran).Q.Bof then Exit;
  {$ENDIF}

  inherited;

  if TypeAction <> taModifEnSerie then begin // VL 02/02/2005
    bStdCegid := EstSpecif('51502');
    {$IFDEF EAGLCLIENT}
    TmpAction := taModif;
    {$ELSE}
    TmpAction := taModifEnSerie;
    {$ENDIF}

    if ((bStdCegid) and (ctxStandard in V_PGI.PGIContexte)) then bStdCegid := True
                                                            else bStdCegid := False;
    {$IFDEF EAGLCLIENT}
      Q := TFMul(Ecran).Q;
      Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);
    {$ELSE}
      Q := TFMul(Ecran).Q;
    {$ENDIF}

    if (((Q.FindField('RB_PREDEFINI').AsString = 'CEG') and (not bStdCegid)) or
      ((Q.FindField('RB_PREDEFINI').AsString = 'STD') and (not (ctxStandard in V_PGI.PGIContexte)))) then
        TmpAction :=taConsult;

    if assigned(Q.FindField('RB_NODOSSIER' )) then
      ParametrageRubrique(Q.FindField('RB_RUBRIQUE' ).AsString+';'+Q.FindField('RB_NODOSSIER' ).AsString+';'+Q.FindField('RB_PREDEFINI' ).AsString+';'+Q.FindField('RB_DATEVALIDITE' ).AsString ,TmpAction, FCtxSurRubrique)
    else  ParametrageRubrique(Q.FindField('RB_RUBRIQUE' ).AsString,TmpAction, FCtxSurRubrique);

    if TmpAction <> taConsult then
      TFMul(Ecran).BChercheClick(nil) ;
   end;

  Screen.Cursor:=SyncrDefault ;
end;

{---------------------------------------------------------------------------------------}
Procedure TOF_CPMULRUB.ChargeCbTypRub ;
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF TRESO}

  {$ELSE}
  if Budget then begin
     RB_TYPERUB.DataType  := 'ttRubTypeBud' ;
     TRB_SIGNERUB.Caption := LMsg.Mess[7] ;
     RB_SIGNERUB.Datatype := 'ttBudJal' ;
     RB_SIGNERUB.Name     := 'RB_BUDJAL' ;
   end
   else
     {Récupère les types de rubriques depuis COMMUN}
     FactoriseComboTypeRub(RB_TYPERUB) ;
  {$ENDIF}
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULRUB.RecupListeRub;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  Fiche : TFMul;
begin
  Fiche := TFMul(Ecran);
  ListeRub:='';
  for n := 0 to Fiche.FListe.nbSelected - 1 do begin
    Fiche.FListe.GotoLeBookmark(n);
    {$IFDEF EAGLCLIENT}
    Fiche.Q.TQ.Seek(Fiche.FListe.Row - 1);
    {$ENDIF}
    ListeRub := ListeRub + Fiche.Q.Findfield('RB_RUBRIQUE').AsString + ',';
    {$IFDEF EAGLCLIENT}
    Fiche.FListe.FlipSelection(n);
    {$ENDIF}
  end ;
  {$IFNDEF EAGLCLIENT}
  Fiche.FListe.FlipSelection;
  {$ENDIF}
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULRUB.EnvoiListeRub ;
{---------------------------------------------------------------------------------------}
var
  TheList : TStringList ;
  n, p    : Integer ;
begin
  TheList := TStringList.Create;
  if ListeRub = '' then Exit;

  while (ListeRub <> '') do TheList.Add(ReadTokenStV(ListeRub)) ;

  for n := 1 to THGrid(TFMul(Ecran).FListe).RowCount - 1 do begin
    for p :=0 to TheList.Count - 1 do
      if TheList[p] = THGrid(TFMul(Ecran).FListe).Cells[1, n] then
        {$IFDEF EAGLCLIENT}
        TFMul(Ecran).FListe.FlipSelection(n);
        {$ENDIF}
  end;
  {$IFNDEF EAGLCLIENT}
  TFMul(Ecran).FListe.FlipSelection;
  {$ENDIF}
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULRUB.BInsertClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
{$IFDEF EAGLCLIENT}
  ParametrageRubrique('',taCreat, FCtxSurRubrique);
{$ELSE}
 ParametrageRubrique('',taCreatEnSerie, FCtxSurRubrique);
{$ENDIF}
  // JP 09/09/2003 : Pour le rafraîchissement de la liste
  TToolbarButton97(GetControl('BCHERCHE')).OnClick(Self);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULRUB.BInsertPersoClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var Q : THQuery;
begin
   {$IFDEF EAGLCLIENT}
      Q := TFMul(Ecran).Q;
      Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);
   {$ELSE}
      Q := TFMul(Ecran).Q;
   {$ENDIF}
//  ParametrageRubrique(GetField('RB_RUBRIQUE'),taCreat, FCtxSurRubrique,False,True);
    if assigned(Q.FindField('RB_NODOSSIER' )) then
      ParametrageRubrique(Q.FindField('RB_RUBRIQUE' ).AsString+';'+Q.FindField('RB_NODOSSIER' ).AsString+';'+Q.FindField('RB_PREDEFINI' ).AsString+';'+Q.FindField('RB_DATEVALIDITE' ).AsString ,taCreat, FCtxSurRubrique,False,True)
    else  ParametrageRubrique(Q.FindField('RB_RUBRIQUE' ).AsString,taCreat, FCtxSurRubrique,False,True);

  // Pour le rafraîchissement de la liste
  TToolbarButton97(GetControl('BCHERCHE')).OnClick(Self);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULRUB.OnLoad ;
{---------------------------------------------------------------------------------------}
begin
  inherited;

  if RubDeRub then EnvoiListeRub ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULRUB.OnClose ;
{---------------------------------------------------------------------------------------}
begin
  FermeMsg ;
  inherited ;
  if RubDeRub then RecupListeRub ;
end ;

{Gestion des message}
{---------------------------------------------------------------------------------------}
procedure TOF_CPMULRUB.InitMsg;
{---------------------------------------------------------------------------------------}
begin
  LMsg:= THMsgBox.create(FMenuG);
  {00}LMsg.Mess.Add('Visualisation des rubriques');
  {01}LMsg.Mess.Add('Modification des rubriques');
  {02}LMsg.Mess.Add('Modification de zones en série des rubriques');
  {03}LMsg.Mess.Add('Modifier en série');
  {04}LMsg.Mess.Add('Visualisation des variables');
  {05}LMsg.Mess.Add('Modification des variables');
  {06}LMsg.Mess.Add('Modification de zones en série des variables');
  {07}LMsg.Mess.Add('&Budget');
  {08}LMsg.Mess.Add('Rubriques');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULRUB.FermeMsg ;
{---------------------------------------------------------------------------------------}
begin
  LMsg.Free ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 17/11/2005
Modifié le ... :   /  /
Description .. : Appelé à chaque changement de ligne dans la grille de
Suite ........ : sélection
Mots clefs ... :
*****************************************************************}
procedure TOF_CPMULRUB.FListeOnRowEnter(Sender: TObject);
begin
  // 
end;

procedure TOF_CPMULRUB.OnDisplay;
begin
  inherited;
  FListeOnRowEnter(nil);
end;

Initialization
  registerclasses ( [ TOF_CPMULRUB ] ) ;
end.



