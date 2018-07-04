{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 17/04/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPSUIVIACC ()
Mots clefs ... : TOF;CPSUIVIACC
*****************************************************************}
Unit CPSUIVIACC_TOF;

Interface

Uses Windows,     // VK_F5
     StdCtrls,
     Controls,
     Classes,
     Grids,       // TGridDrawState
     Graphics,    // clRed
{$IFDEF EAGLCLIENT}
     MaineAGL,    // AGLLanceFiche
     eMul,        // TFMul
{$ELSE}
     db,
     dbtables,
     FE_Main,     // AGLLanceFiche
     Mul,         // TFMul
     HDB,         // THDBGrid
{$ENDIF}
     Saisie,      // TrouveEtLanceSaisie
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     Ent1,        // tProfilTraitement, ExoToDates, LibellesTableLibre, WhereProfilUser
     HTB97,       // TToolBarButton97
     HQry,        // RecupWhereCritere
     TofVerifRib, // CPLanceFiche_VerifRib
     MulSMPUtil,  // ModifRibSurMul
     SaisUtil,    // QuelExo
     SaisBor,     // LanceSaisieFolio
     Choix,       // Choisir
     UTOF,
     ParamSoc,		// GetParamSocSecur YMO
     Lookup;  //fb 19/06/2006 FQ12487

procedure SuivAcceptation(Qui : tProfilTraitement);

Type
  TOF_CPSUIVIACC = Class (TOF)
    HM: THMsgBox;
{$IFDEF EAGLCLIENT}
    FListe : THGrid;
{$ELSE}
    FListe : THDBGrid;
{$ENDIF}
    Q : THQuery;
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
    procedure BCtrlRibClick(Sender: TObject);
    procedure FLettrageChange(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure BChercheClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject);
    procedure bSelectAllClick(Sender: TObject);
{$IFDEF EAGLCLIENT}
    procedure FListeRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure FListeDrawCell(Sender: TObject; ACol, ARow: Integer; Rect: TRect; State: TGridDrawState);
{$ELSE}
    procedure FListeRowEnter(Sender: TObject);
    procedure FListeDrawDataCell(Sender: TObject; const Rect: TRect; Field: TField; State: TGridDrawState);
{$ENDIF}
{b fb 19/06/2006 FQ12487}
    procedure SelectionElipsisClick(Sender : TObject);
    procedure AuxiliaireElipsisClick(Sender : TObject);
{e fb 19/06/2006 FQ12487}
    procedure AuxiElipsisClick(Sender : TObject);
  private
    GeneOpe : String ;
    Qui : tProfilTraitement ;
    Selection, Auxiliaire : THCritMaskEdit;  {fb 19/06/2006 FQ12487}
    procedure InitCriteres ;
    procedure QueLesLCR ;
    procedure ValideAcc ;
    procedure ClickModifRib ;
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
  , UTofMulParamGen; {13/04/07 YMO F5 sur Auxiliaire }


procedure SuivAcceptation(Qui : tProfilTraitement);
var
  szArg : String;
begin
 if Qui = prClient then szArg := '1' else
 if Qui = prFournisseur then szArg := '2' else
 if Qui = prEtablissement then szArg := '3' else
 szArg := '4';

 AGLLanceFiche('CP','CPSUIVIACC','','', szArg);
end;

procedure TOF_CPSUIVIACC.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPSUIVIACC.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPSUIVIACC.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPSUIVIACC.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CPSUIVIACC.OnArgument (S : String ) ;
begin
  Inherited ;
  Q := TFMul(Ecran).Q;

  // Création des contrôles
  HM := THMsgBox.create(FMenuG);
  InitMsgBox;

  // Récupère les arguments
  if S = '1' then Qui := prClient else
  if S = '2' then Qui := prFournisseur else
  if S = '3' then Qui := prEtablissement else
  Qui := prAucun;
  TFMul(Ecran).FNomFiltre := 'CPSUIVACC';

  // Evénements des contrôles
{$IFDEF EAGLCLIENT}
  FListe := THGrid(TFMul(Ecran).FListe);
//  FListe.OnDrawCell := FListeDrawCell;
{$ELSE}
  FListe := THDBGrid(TFMul(Ecran).FListe);
//  FListe.OnDrawDataCell := FListeDrawDataCell;
{$ENDIF}
  THValComboBox(GetControl('FLETTRAGE',True)).OnChange := FLettrageChange;
  THValComboBox(GetControl('E_EXERCICE',True)).OnChange := E_EXERCICEChange;
  TToolBarButton97(GetControl('BCTRLRIB',True)).OnClick := BCtrlRibClick;
  TToolBarButton97(GetControl('BOUVRIR',True)).OnClick := BOuvrirClick;
  TToolBarButton97(GetControl('BCHERCHE',True)).OnClick := BChercheClick;
  TToolBarButton97(GetControl('BSELECTALL',True)).OnClick := bSelectAllClick;  
  Ecran.OnKeyDown := FormKeyDown;
  FListe.OnRowEnter := FListeRowEnter;

{b fb 19/06/2006 FQ12487}
  Selection       := THCritMaskEdit(GetControl('E_GENERAL', true));
  Auxiliaire      := THCritMaskEdit(GetControl('E_AUXILIAIRE', true));
  Selection.OnElipsisClick := SelectionElipsisClick;
  if GetParamSocSecur('SO_CPMULTIERS', false) then
    Auxiliaire.OnElipsisClick := AuxiElipsisClick
  else
    Auxiliaire.OnElipsisClick := AuxiliaireElipsisClick;
{e fb 19/06/2006 FQ12487}
  // FormCreate
  TFMul(Ecran).MemoStyle := msBook;

  // FormShow
  Ecran.HelpContext := 3710500;
  InitCriteres;
  QuelesLCR;
  SetControlText('FLETTRAGE', 'NL');
  SetControlVisible('FLETTRAGE', False);
  SetControlVisible('TFLETTRAGE', False);
  SetControlText('XX_WHEREVIDE', '');
end ;

procedure TOF_CPSUIVIACC.OnClose ;
begin
  Inherited ;
  HM.Free;
end ;

procedure TOF_CPSUIVIACC.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPSUIVIACC.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_CPSUIVIACC.InitMsgBox;
begin
  HM.Mess.Add('Opération à effectuer');
  HM.Mess.Add('ATTENTION : Certaines échéances en cours de traitement n''ont pas été mises à jour.');
  HM.Mess.Add('2;?caption?;Confirmez-vous le traitement ?;Q;YN;Y;Y;');
end;

procedure TOF_CPSUIVIACC.BOuvrirClick(Sender: TObject);
var
  Ope : String;
  i   : integer;
begin
  GeneOpe := '';
  // Opération à effectuer
  Ope := Choisir(HM.Mess[0],'COMMUN','CO_LIBELLE','CO_CODE','CO_TYPE="TLC" AND CO_CODE<>"NON" AND CO_CODE<>"BOR"','');
  if (Ope = '') then Exit
  else begin
    // Confirmez-vous le traitement ?
    if HM.Execute(2,Ecran.Caption,'')<>mrYes then Exit;
    GeneOpe := Ope;
  end;

  if Not FListe.AllSelected then begin
    for i := 0 to FListe.NbSelected-1 do begin
      FListe.GotoLeBookmark(i);
      {$IFDEF EAGLCLIENT}
      TFMul(Ecran).Q.TQ.Seek(FListe.Row-1);
      {$ENDIF}
      if Transactions(ValideAcc,1)<>oeOk then begin
        // Certaines échéances en cours de traitement n''ont pas été mises à jour.
        MessageAlerte(HM.Mess[1]);
        Break;
      end;
    end;
    end
  else begin
    Q.First;
    While Not Q.EOF do begin
      if Transactions(ValideAcc,1)<>oeOk then begin
        // Certaines échéances en cours de traitement n''ont pas été mises à jour.
        MessageAlerte(HM.Mess[1]);
        Break;
      end;
      Q.Next;
    end;
  end;
  BChercheClick(Nil);
end;

procedure TOF_CPSUIVIACC.BCtrlRibClick(Sender: TObject);
var
  StWRib : String ;
  i : Integer;
begin
  inherited;
  StWRib := RecupWhereCritere(TFMul(Ecran).Pages) ;
  if (StWRib = '') then Exit;

  // Si on n'est pas en Sélection inversée
  if ((Not FListe.AllSelected) and (FListe.NbSelected>0) and (FListe.NbSelected<100)) then begin
    // Si on n'a pas tous sélectionné ET qu'il y a au moins 1 et 100 au plus lignes sélectionnées
    // Rajoute une clause au WHERE
    StWRib := StWRib+' AND (';
    for i:=0 to FListe.NbSelected-1 do begin
      FListe.GotoLeBookmark(i) ;
{$IFDEF EAGLCLIENT}
      TFMul(Ecran).Q.TQ.Seek(FListe.Row-1);
{$ENDIF}
      StWRib := StWRib + ' (E_NUMEROPIECE='+ Q.FindField('E_NUMEROPIECE').AsString +' AND E_NUMLIGNE='+ Q.FindField('E_NUMLIGNE').AsString +' AND E_JOURNAL="'+ Q.FindField('E_JOURNAL').AsString +'") OR';
    end;
    // Efface le dernier OR et rajoute ')'
    System.Delete(StWRib,length(StWRib)-2,3);
    StWRib := StWRib +')';
  end;
  If StWRib<>'' Then CPLanceFiche_VerifRib('WHERE='+StWRib);
end;

procedure TOF_CPSUIVIACC.ClickModifRib;
var
  iNumLigne : Integer ;
{$IFNDEF EAGLCLIENT}
  RJal,RExo : String ;
  RDate : TDateTime ;
  RNumP,RNumL,RNumEche : Integer ;
{$ENDIF}
begin
  iNumLigne := FListe.Row;
{$IFDEF EAGLCLIENT}
  Q.TQ.Seek(FListe.Row-1);
  If ModifRibSurMul(Q.TQ,FALSE,TRUE) then begin
{$ELSE}
  If ModifRibSurMul(Q,FALSE,TRUE) then begin
{$ENDIF}
    Application.ProcessMessages;
    BChercheClick(Nil);
    FListe.Row := iNumLigne;
{$IFDEF EAGLCLIENT}
    Q.TQ.Seek(TFMul(Ecran).FListe.Row-1);
{$ELSE}
    // Se repositionne sur l'enregistrement
    RJal := Q.FindField('E_JOURNAL').AsString;
    RExo := QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime));
    RDate := Q.FindField('E_DATECOMPTABLE').AsDateTime;
    RNumP := Q.FindField('E_NUMEROPIECE').AsInteger;
    RNumL := Q.FindField('E_NUMLIGNE').AsInteger;
    RNumEche := Q.FindField('E_NUMECHE').AsInteger;
    Q.Locate('E_JOURNAL;E_EXERCICE;E_DATECOMPTABLE;E_QUALIFPIECE;E_NUMEROPIECE;E_NUMLIGNE;E_NUMECHE',
             VarArrayOf([RJal,RExo,RDate,'N',RNumP,RNumL,RNumEche]),[]);
{$ENDIF}
  end;
end;

procedure TOF_CPSUIVIACC.E_EXERCICEChange(Sender: TObject);
begin
  ExoToDates(GetControlText('E_EXERCICE'), GetControl('E_DATECOMPTABLE',True), GetControl('E_DATECOMPTABLE_',True) );
end;

procedure TOF_CPSUIVIACC.FLettrageChange(Sender: TObject);
begin
  if GetControlText('FLETTRAGE') = 'LE' then SetControlText('XX_WHERELETTRAGE', 'E_LETTRAGE<>""') else
  if GetControlText('FLETTRAGE') = 'NL' then SetControlText('XX_WHERELETTRAGE', '(E_ETATLETTRAGE="AL" OR E_ETATLETTRAGE="PL")') else
  SetControlText('XX_WHERELETTRAGE', '');
end;

procedure TOF_CPSUIVIACC.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  Case Key of
    VK_F5 : begin Key:=0; ClickModifRib; end;
  end;
end;

procedure TOF_CPSUIVIACC.BChercheClick(Sender: TObject);
{$IFDEF EAGLCLIENT}
var
  Cancel : Boolean;
{$ENDIF}
begin
  TFMul(Ecran).BChercheClick(Sender); // inherited;
{$IFDEF EAGLCLIENT}
  Cancel := False;
  FListeRowEnter(nil,FListe.Row,Cancel,False);
{$ELSE}
  FListeRowEnter(nil);
{$ENDIF}
end;

procedure TOF_CPSUIVIACC.InitCriteres;
var
  St : String ;
begin
  LibellesTableLibre(TTabSheet(GetControl('PZLIBRE', True)),'TE_TABLE','E_TABLE','E') ;
  if VH^.CPExoRef.Code<>'' then begin
    SetControlText('E_EXERCICE', VH^.CPExoRef.Code);
    SetControlText('E_DATECOMPTABLE', DateToStr(VH^.CPExoRef.Deb));
    SetControlText('E_DATECOMPTABLE_', DateToStr(VH^.CPExoRef.Fin));
    end
  else begin
    SetControlText('E_EXERCICE', VH^.Entree.Code);
    SetControlText('E_DATECOMPTABLE', DateToStr(VH^.Entree.Deb));
    SetControlText('E_DATECOMPTABLE_', DateToStr(VH^.Entree.Fin));
  end;
  SetControlText('E_DATEECHEANCE', StDate1900);
  SetControlText('E_DATEECHEANCE_', StDate2099);
  SetControlText('E_DEVISE', V_PGI.DevisePivot);
  GeneOpe := '';
  SetControlText('XX_WHEREPROFIL', '');
{$IFDEF EAGLCLIENT}
  St := WhereProfilUser(Q.TQ,Qui);
{$ELSE}
  St := WhereProfilUser(Q,Qui);
{$ENDIF}
  SetControlText('XX_WHEREPROFIL', St);
end;

procedure TOF_CPSUIVIACC.QueLesLCR;
var
  St,sMode,sCat,sAcc : String ;
  i : integer ;
begin
  SetControlText('XX_WHEREMP', '');
  for i := 0 to VH^.MPACC.Count-1 do begin
    St := VH^.MPACC[i];
    sMode := ReadtokenSt(St);
    sAcc := ReadtokenSt(St);
    sCat := ReadtokenSt(St);
    if sCat<>'LCR' then Continue ;
    if ((sAcc<>'BOR') and (sAcc<>'NON')) then SetControlText('XX_WHEREMP', GetControlText('XX_WHEREMP') + ' OR E_MODEPAIE="'+sMode+'"');
  end;

  if GetControlText('XX_WHEREMP') = '' then begin
    SetControlText('XX_WHEREMP', 'E_MODEPAIE="aaa"');
    end
  else begin // ne rien trouver
    St := GetControlText('XX_WHEREMP');
    System.Delete(St,1,4);
    SetControlText('XX_WHEREMP', St);
  end;
end;

procedure TOF_CPSUIVIACC.ValideAcc;
var
  SW : String ;
  Nb     : integer ;
begin
  SW := 'E_JOURNAL="'+Q.FindField('E_JOURNAL').AsString+'"'
       +' AND E_EXERCICE="'+QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime))+'"'
       +' AND E_DATECOMPTABLE="'+USDATETIME(Q.FindField('E_DATECOMPTABLE').AsDateTime)+'"'
       +' AND E_NUMEROPIECE='+Q.FindField('E_NUMEROPIECE').AsString
       +' AND E_NUMLIGNE='+Q.FindField('E_NUMLIGNE').AsString
       +' AND E_NUMLIGNE='+Q.FindField('E_NUMLIGNE').AsString
       +' AND E_NUMECHE='+Q.FindField('E_NUMECHE').AsString
       +' AND E_QUALIFPIECE="N" ';
  Nb := ExecuteSQL('UPDATE ECRITURE SET E_CODEACCEPT="'+GeneOpe+'" WHERE '+SW);
  if Nb<>1 then V_PGI.IoError:=oeUnknown ;
end;

{$IFDEF EAGLCLIENT}
procedure TOF_CPSUIVIACC.FListeRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  TFMul(Ecran).FListeRowEnter(Sender, Ou, Cancel, Chg);  // inherited;
{$ELSE}
procedure TOF_CPSUIVIACC.FListeRowEnter(Sender: TObject);
begin
//  TFMul(Ecran).FListeRowEnter(Sender);  // inherited;
{$ENDIF}
  VH^.MPModifFaite := False;
  If Q.FindField('E_DATECOMPTABLE')<>NIL then VH^.MPPop.MPExoPop  := QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime)) ;
  If Q.FindField('E_GENERAL')<>NIL       then VH^.MPPop.MPGenPop  := Q.FindField('E_GENERAL').AsString ;
  If Q.FindField('E_AUXILIAIRE')<>NIL    then VH^.MPPop.MPAuxPop  := Q.FindField('E_AUXILIAIRE').AsString ;
  If Q.FindField('E_JOURNAL')<>NIL       then VH^.MPPop.MPJalPop  := Q.FindField('E_JOURNAL').AsString ;
  If Q.FindField('E_NUMEROPIECE')<>NIL   then VH^.MPPop.MPNumPop  := Q.FindField('E_NUMEROPIECE').AsInteger ;
  If Q.FindField('E_NUMLIGNE')<>NIL      then VH^.MPPop.MPNumLPop := Q.FindField('E_NUMLIGNE').AsInteger ;
  If Q.FindField('E_NUMECHE')<>NIL       then VH^.MPPop.MPNumEPop := Q.FindField('E_NUMECHE').AsInteger ;
  If Q.FindField('E_DATECOMPTABLE')<>NIL then VH^.MPPop.MPDatePop := Q.FindField('E_DATECOMPTABLE').AsDateTime ;
end;

{$IFDEF EAGLCLIENT}
procedure TOF_CPSUIVIACC.FListeDrawCell(Sender: TObject; ACol, ARow: Integer; Rect: TRect; State: TGridDrawState);
var
  szRIB : String;
begin
  if ((Q.TQ.EOF) and (Q.TQ.BOF)) then Exit ;
  if ((Q.TQ.GetNomChamp(ACol) ='E_AUXILIAIRE') and (Q.FindField('E_RIB')<>Nil)) then begin
    szRIB := Q.FindField('E_RIB').AsString;
    if Trim(szRib)<>'' then Exit ;
    FListe.Canvas.Brush.Color := clRed;
    FListe.Canvas.Brush.Style := bsSolid;
    FListe.Canvas.Pen.Color := clRed;
    FListe.Canvas.Pen.Mode := pmCopy;
    FListe.Canvas.Pen.Width := 1;
    FListe.Canvas.Rectangle(Rect.Right-5,Rect.Top+1,Rect.Right-1,Rect.Top+5);
  end;
end;
{$ELSE}
procedure TOF_CPSUIVIACC.FListeDrawDataCell(Sender: TObject; const Rect: TRect; Field: TField; State: TGridDrawState);
var
  szRIB : String;
begin
  if ((Q.EOF) and (Q.BOF)) then Exit ;
  if ((Field.FieldName='E_AUXILIAIRE') and (Q.FindField('E_RIB')<>Nil)) then begin
    szRIB := Q.FindField('E_RIB').AsString;
    if Trim(szRib)<>'' then Exit ;
    FListe.Canvas.Brush.Color := clRed;
    FListe.Canvas.Brush.Style := bsSolid;
    FListe.Canvas.Pen.Color := clRed;
    FListe.Canvas.Pen.Mode := pmCopy;
    FListe.Canvas.Pen.Width := 1;
    FListe.Canvas.Rectangle(Rect.Right-5,Rect.Top+1,Rect.Right-1,Rect.Top+5);
  end;
end;
{$ENDIF}

procedure TOF_CPSUIVIACC.FListeDblClick(Sender: TObject);
var
  sMode : String ;
begin
  TFMul(Ecran).FListeDblClick(Sender); // inherited;

{$IFDEF EAGLCLIENT}
  if ((Q.TQ.EOF) and (Q.TQ.BOF)) then Exit ;
{$ELSE}
  if ((Q.EOF) and (Q.BOF)) then Exit ;
{$ENDIF}
  sMode := Q.FindField('E_MODESAISIE').AsString;
{$IFDEF EAGLCLIENT}
  if ((sMode<>'') and (sMode<>'-')) then LanceSaisieFolio(Q.TQ,TFMul(Ecran).TypeAction)
                                    else TrouveEtLanceSaisie(Q.TQ,taConsult,'N');
{$ELSE}
  if ((sMode<>'') and (sMode<>'-')) then LanceSaisieFolio(Q,TFMul(Ecran).TypeAction)
                                    else TrouveEtLanceSaisie(Q,taConsult,'N');
{$ENDIF}
end;

procedure TOF_CPSUIVIACC.bSelectAllClick(Sender: TObject);
{$IFDEF EAGLCLIENT}
var
  Fiche : TFMul;
{$ENDIF}
begin
{$IFDEF EAGLCLIENT}
Fiche := TFMul(Ecran);
if Fiche.bSelectAll.Down then
begin
  if not Fiche.FetchLesTous then
  begin
    Fiche.bSelectAllClick(nil);
    Fiche.bSelectAll.Down := False;
    exit;
  end;
end;
{$ENDIF}
TFMul(Ecran).bSelectAllClick(nil);
end;

{b fb 19/06/2006 FQ12487}
procedure TOF_CPSUIVIACC.SelectionElipsisClick(Sender : TObject);
begin
  LookUpList(TControl(Sender),'Recherche d''un compte collectif','GENERAUX','G_GENERAL','G_LIBELLE',
  '(G_SUIVITRESO="ENC" OR G_SUIVITRESO="MIX") AND (G_COLLECTIF="X" OR G_LETTRABLE="X")','G_GENERAL',true,-1);
end;

procedure TOF_CPSUIVIACC.AuxiliaireElipsisClick(Sender : TObject);
begin
  LookUpList(TControl(Sender),'Recherche d''un compte auxiliaire','TIERS','T_AUXILIAIRE','T_LIBELLE',
  '(T_NATUREAUXI="AUD" OR T_NATUREAUXI="CLI" OR T_NATUREAUXI="DIV")','T_AUXILIAIRE',true,-1)
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 13/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CPSUIVIACC.AuxiElipsisClick( Sender : TObject );
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;

{e fb 19/06/2006 FQ12487}
Initialization
  registerclasses ( [ TOF_CPSUIVIACC ] ) ;
end.

