unit YYMULEXPERTSCAN_TOF;

Interface

Uses
     Classes,
     UTob,
{$IFDEF EAGLCLIENT}
     MainEAgl,        // AGLLanceFiche
     UtileAgl, {LanceEtatTob}
{$ELSE}
     Fe_Main,         // AGLLanceFiche
     Db,
//     EdtREtat, {LanceEtatTob}
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     ULibEcriture ,
     ULibExpertScan,
     uTofViergeMul;   // Fiche Ancetre Vierge MUL

type

  TOF_YYMULEXPERTSCAN = Class (TOF_ViergeMul)
   procedure InitPopUp ;
   procedure OnClickGenererArbo ( Sender : TObject) ;
   procedure OnClickHPUP ( Sender : TObject) ;
   procedure OnClickHPDOWN ( Sender : TObject) ;
   procedure OnClickGenererXML ( Sender : TObject) ;
   procedure OnClickSynchroEws ( Sender : TObject) ;
   procedure OnClickRecuperEws ( Sender : TObject) ;
   procedure OnClickDebloquerEws ( Sender : TObject) ;
   procedure OnClickActiverScanEws ( Sender : TObject) ;
   procedure OnPopUpPopF11( Sender : TObject ) ; override ;
   procedure OnClickBCherche (Sender : TObject) ; override ;
   private
     FExpertScan : TExpertScan ;
     FTOBDossier : TOB ;
   protected
    procedure RemplitATobFListe ; override ;
    procedure InitControl ; override ;
    function  AjouteATobFListe( vTob : Tob) : Boolean; override ;
    procedure OnErrorScan (sender : TObject; Error : TRecErrorScan ) ;
    procedure OnProgress ( vStDossier,vStNature, vStMessage : string ) ;
    procedure OnInfo ( vStMessage : string ) ;
   public
    procedure OnLoad                   ; override ;
    procedure OnArgument              (S : String ) ; override ;
    procedure OnClose ; override ;
  end ;

procedure YYLanceFiche_MULEXPERTSCAN ;


implementation

uses
 HSysMenu,     // HSystemMenu
 UtileWS ,
 HEnt1 ,
 ed_tools ,
 Controls ,
 SaisComm , // pour le GetO
 hmsgbox ,
 StdCtrls , // pour le cbChecked
 paramSoc ,
 CbpTrace ,
 Forms ,
 Menus,  // PopUpMenu
 SysUtils ;

const
 cFI_TABLE    = 'YYEXPERTSCAN' ;

{ TOF_CPMULSCAN }

procedure YYLanceFiche_MULEXPERTSCAN ;
begin
 AGLLanceFiche('YY','YYEXPERTSCAN', '', '', '') ;
end ;

function TOF_YYMULEXPERTSCAN.AjouteATobFListe(vTob: Tob) : Boolean ;
var
 lTOBDoss : TOB ;
begin
 if vTOB.GetString('BOSCANEWS') = '' then
  vTOB.PutValue('BOSCANEWS' , '-') ;
 if vTOB.GetString('BOHP') = '' then
  vTOB.PutValue('BOHP' , '-') ;
 if vTOB.GetString('BOHP') <> '-' then
  vTOB.AddChampSupValeur('NBFICHHP',FExpertScan.ScanNbFichierD(vTOB.GetString('DOS_NODOSSIER') ) )
   else
     vTOB.AddChampSupValeur('NBFICHHP', 0 ) ;
 lTOBDoss := FTOBDossier.FindFirst(['DOSSIER'],[vTOB.GetString('dos_nodossier')], true) ;
 vTOB.AddChampSupValeur('BOVERROUEWS' , '-' ) ;
 if lTOBDoss <> nil then
  begin
   vTOB.AddChampSupValeur('NBFICHEWS' , lTOBDoss.GetValue('NBFICHEWS') ) ;
   vTOB.PutValue('BOVERROUEWS' , lTOBDoss.GetValue('BOVERROUEWS') ) ;
  end ;
 result := true ;
end;


procedure TOF_YYMULEXPERTSCAN.InitControl;
begin

 inherited;

 PageControl.ActivePage := PageControl.Pages[0] ;
 FListe.ColTypes[3]     := 'B' ;
 FListe.ColTypes[4]     := 'B' ;
 FListe.ColTypes[5]     := 'B' ;

end;

procedure TOF_YYMULEXPERTSCAN.OnArgument(S: String);
begin

 Trace.TraceError('EXPERTSCAN','TOF_YYMULEXPERTSCAN.OnArgument') ;

 FFI_Table       := cFI_TABLE;
 FStListeParam   := 'YYEXPERTSCAN' ;

 inherited ;

 try
  FExpertScan                := TExpertScan.Create ;
  FExpertScan.OnError        := OnErrorScan ;
  FExpertScan.Path           := GetParamsocDpSecur('so_cpscanpath','') ;
  FExpertScan.UserPGI        := V_PGI.UserLogin ;
  FExpertScan.PwUserPGI      := V_PGI.Password ;
  FExpertScan.OnProgress     := OnProgress ;
  FExpertScan.OnInfo         := OnInfo ;
  SourisSablier ;
  InitMoveProgressForm(nil, '',traduireMemoire('Test d''eWS en cours...'), 10, False, True);
  try
  if EwsLibrairieOk(false) then
   begin
    MoveCurProgressForm(traduireMemoire('eWS Présent, création des objects Expert Scan') );
    FExpertScan.eWSSecurite    := GetParamsocDpSecur('SO_NECABIDENT', '') ;
    FExpertScan.EwsVersion     := GetEwsVersion ;
    MoveCurProgressForm(traduireMemoire('eWS Présent, création des objects de liaison eWS') );
    FExpertScan.V_eWS          := V_eWS ;
    FExpertScan.PwSystem       := GetParamsocDpSecur('SO_CPSCANPW', '') ;
    FExpertScan.UserSystem     := GetParamsocDpSecur('SO_CPSCANUSER', '') ;
    FExpertScan.Ip             := GetParamsocDpSecur('SO_CPSCANIP', '') ;
    FExpertScan.CWASName       := GetParamsocDpSecur('SO_CWS', '') ;
   end ;
  finally
   FiniMoveProgressForm ;
   SourisNormale ;
  end ;
 except
  on e:exception do
   pgiError(e.Message) ;
 end ;
 
 FTOBDossier := TOB.Create('',nil,-1) ;
 InitPopUp ;
 OnPopUpPopF11(nil) ;

 Ecran.PopupMenu := PopF11 ;

end;

procedure TOF_YYMULEXPERTSCAN.OnClose;
begin
 Trace.TraceError('EXPERTSCAN','TOF_YYMULEXPERTSCAN.OnClose Debut') ;
 FTOBDossier.Free ;
 FExpertScan.Free ;
 Trace.TraceError('EXPERTSCAN','TOF_YYMULEXPERTSCAN.OnClose Fin') ;
 inherited;
end;

procedure TOF_YYMULEXPERTSCAN.OnLoad;
begin
  inherited;
end ;


procedure TOF_YYMULEXPERTSCAN.InitPopUp ;
var
 lPopUp : TPopUpMenu ;
 i      : integer ;
begin

 if csDestroying in Ecran.ComponentState then Exit ;

 Trace.TraceError('EXPERTSCAN','TOF_YYMULEXPERTSCAN.InitPopUp');

 lPopUp := TPopUpMenu(GetControl('PopUpOutil', True)) ;

 for i := 0 to lPopUp.Items.Count - 1 do
  begin

   if lPopUp.Items[i].Name = 'GENERERARBO' then
    begin
     lPopUp.Items[i].OnClick := OnClickGenererArbo ;
     Continue;
    end; // if

   if lPopUp.Items[i].Name = 'HPUP' then
    begin
     lPopUp.Items[i].OnClick := OnClickHPUP ;
     Continue;
    end; // if

   if lPopUp.Items[i].Name = 'HPDOWN' then
    begin
     lPopUp.Items[i].OnClick := OnClickHPDOWN ;
     Continue;
    end; // if

   if lPopUp.Items[i].Name = 'GENERERXML' then
    begin
     lPopUp.Items[i].OnClick := OnClickGenererXML ;
     Continue;
    end; // if

   if lPopUp.Items[i].Name = 'SYNCHROEWS' then
    begin
     lPopUp.Items[i].OnClick := OnClickSynchroEws ;
     Continue;
    end; // if

   if lPopUp.Items[i].Name = 'RECUPEREWS' then
    begin
     lPopUp.Items[i].OnClick := OnClickRecuperEws ;
     Continue;
    end; // if

   if lPopUp.Items[i].Name = 'DEBLOQUEREWS' then
    begin
     lPopUp.Items[i].OnClick := OnClickDebloquerEws ;
     Continue;
    end; // if

   if lPopUp.Items[i].Name = 'ACTIVEREWS' then
    begin
     lPopUp.Items[i].OnClick := OnClickActiverScanEws ;
     Continue;
    end; // if

  end ; // for

end ;


procedure TOF_YYMULEXPERTSCAN.OnClickHPUP ( Sender : TObject) ;
var
 lTOB : TOB ;
 i    : integer ;
begin

 if PGIask('Confirmez-vous l''activation du MFP3035 ? ' , 'Activation') <> mrYes then exit ;

 if ABoMultiSelected then
  begin
   for i := 1 to FListe.RowCount - 1 do
    if FListe.IsSelected(i) then
     begin
      lTob := GetO(FListe,i) ;
      CSetParamScanHP(lTOB.GetString('dos_nodossier') ,'X' ) ;
     end ; // if
  end // if
   else
    begin
     lTob := GetO(FListe) ;
     CSetParamScanHP(lTOB.GetString('dos_nodossier') ,'X' ) ;
    end ;

 RefreshPclPge ;

end ;

procedure TOF_YYMULEXPERTSCAN.OnClickHPDOWN ( Sender : TObject) ;
var
 lTOB : TOB ;
 i    : integer ;
begin

 if PGIask('Confirmez-vous la désactivation du MFP3035 ? ' , 'Désactivation') <> mrYes then exit ;

 if ABoMultiSelected then
  begin
   for i := 1 to FListe.RowCount - 1 do
    if FListe.IsSelected(i) then
     begin
      lTob := GetO(FListe,i) ;
      CSetParamScanHP(lTOB.GetString('dos_nodossier') , '-' ) ;
     end ; // if
  end // if
   else
    begin
     lTob := GetO(FListe) ;
     CSetParamScanHP(lTOB.GetString('dos_nodossier') , '-' ) ;
    end ;

 RefreshPclPge ;

end ;

procedure TOF_YYMULEXPERTSCAN.OnClickGenererXML ( Sender : TObject) ;
var
 lTOB : TOB ;
 i    : integer ;
begin

 if PGIask('Confirmez-vous la génération du fichier de configuration du scanner ? ' , 'Génération') <> mrYes then exit ;

 FListe.BeginUpdate ;
 FExpertScan.VideDossier ;

 try

 for i := 1 to FListe.RowCount - 1 do
  begin
   lTob := GetO(FListe,i) ;
   if lTOB.GetString('BOHP') = 'X' then
    FExpertScan.AjouteDossier(lTOB.GetString('dos_nodossier') ) ;
  end ; // for

 FExpertScan.GenererXML ;

 finally
  FListe.EndUpdate ;
  RefreshPclPge ;
 end ;

end ;

procedure TOF_YYMULEXPERTSCAN.OnClickSynchroEws ( Sender : TObject) ;
var
 i : integer ;
 lTOB : TOB ;

 procedure _Synchro ;
 var
  lTOBDoss : TOB ;
 begin
  lTOBDoss := FTOBDossier.FindFirst(['DOSSIER'],[lTOB.GetString('dos_nodossier')], true) ;
  if lTOBDoss = nil then
   begin
    lTOBDoss := TOB.Create('', FTOBDossier, - 1 ) ;
    lTOBDoss.AddChampSupValeur('DOSSIER',lTOB.GetString('dos_nodossier')) ;
    lTOBDoss.AddChampSup('NBFICHEWS',false) ;
    lTOBDoss.AddChampSupValeur('BOVERROUEWS' , '-' ) ;
   end ;
   lTOBDoss.PutValue('NBFICHEWS',FExpertScan.eWSSynchro(lTOB.GetString('dos_nodossier') ) ) ;
   if FExpertScan.eWSEstBloque(lTOB.GetString('dos_nodossier')) then
    lTOBDoss.PutValue('BOVERROUEWS','X') ;
 end ;


begin

 if PGIask('Confirmez-vous la synchronisation eWS ? ' , 'Synchronisation') <> mrYes then exit ;

 SourisSablier ;
 InitMoveProgressForm(nil, '',traduireMemoire('Synchro eWS en cours...'), FListe.RowCount - 1 , False, True);

 try
 try

 if ABoMultiSelected then
  begin
   for i := 1 to FListe.RowCount - 1 do
    if FListe.IsSelected(i) then
     begin
      lTOB := GetO(FListe,i) ;
      MoveCurProgressForm(traduireMemoire('Mise à jour du dossier ') + lTOB.GetString('dos_nodossier') );
      _Synchro ;
     end ; // if
  end // if
   else
    begin
     lTOB := GetO(FListe) ;
     MoveCurProgressForm(traduireMemoire('Mise à jour du dossier ') + lTOB.GetString('dos_nodossier') );
     _Synchro ;
    end ;

 RefreshPclPge ;

 finally
  FiniMoveProgressForm ;
  SourisNormale ;
 end ;

 except
  on e:exception do
   PGIError('Erreur lors de la synchronisation' + #10#13  + e.Message) ;
 end ;

end ;

procedure TOF_YYMULEXPERTSCAN.OnClickRecuperEws ( Sender : TObject) ;
var
 i : integer ;
 lTOB : TOB ;
begin

 if PGIask('Confirmez-vous la récupération des fichiers eWS ? ' , 'Récupération') <> mrYes then exit ;

 SourisSablier ;
 InitMoveProgressForm(nil, '',traduireMemoire('Récupération eWS en cours...'), FListe.RowCount - 1 , False, True) ;

 try

  if ABoMultiSelected then
  begin
   for i := 1 to FListe.RowCount - 1 do
    if FListe.IsSelected(i) then
     begin
      lTOB := GetO(FListe,i) ;
      MoveCurProgressForm(traduireMemoire('Récupération du dossier ') + lTOB.GetString('dos_nodossier') );
      FExpertScan.eWSRecuperer(lTOB.GetString('dos_nodossier')) ;
     end ; // if
  end // if
   else
    begin
     lTOB := GetO(FListe) ;
     MoveCurProgressForm(traduireMemoire('Récupération du dossier ') + lTOB.GetString('dos_nodossier') );
     FExpertScan.eWSRecuperer(lTOB.GetString('dos_nodossier')) ;
    end ;


 finally
  FiniMoveProgressForm ;
  SourisNormale ;
 end ;

end ;

procedure TOF_YYMULEXPERTSCAN.OnClickDebloquerEws ( Sender : TObject) ;
var
 i : integer ;
 lTOB : TOB ;
begin

 if PGIask('Confirmez-vous le débloquage des fichiers selectionné ? ' , 'Déblocaque') <> mrYes then exit ;

 SourisSablier ;
 InitMoveProgressForm(nil, '',traduireMemoire('Débloquage eWS en cours...'), FListe.RowCount - 1 , False, True) ;

 try

  if ABoMultiSelected then
  begin
   for i := 1 to FListe.RowCount - 1 do
    if FListe.IsSelected(i) then
     begin
      lTOB := GetO(FListe,i) ;
      MoveCurProgressForm(traduireMemoire('Récupération du dossier ') + lTOB.GetString('dos_nodossier') );
      FExpertScan.eWSDebloquer(lTOB.GetString('dos_nodossier')) ;
     end ; // if
  end // if
   else
    begin
     lTOB := GetO(FListe) ;
     MoveCurProgressForm(traduireMemoire('Récupération du dossier ') + lTOB.GetString('dos_nodossier') );
     FExpertScan.eWSDebloquer(lTOB.GetString('dos_nodossier')) ;
    end ;


 finally
  FiniMoveProgressForm ;
  SourisNormale ;
 end ;

end ;

procedure TOF_YYMULEXPERTSCAN.OnClickActiverScanEws ( Sender : TObject) ;
var
 i : integer ;
 lTOB : TOB ;
begin

 if PGIask('Confirmez-vous l''activation du scan Ews sur les fichiers selectionnées ? ' , 'Activation') <> mrYes then exit ;

 SourisSablier ;
 InitMoveProgressForm(nil, '',traduireMemoire('Activation du scan eWS en cours...'), FListe.RowCount - 1 , False, True) ;

 try

  if ABoMultiSelected then
  begin
   for i := 1 to FListe.RowCount - 1 do
    if FListe.IsSelected(i) then
     begin
      lTOB := GetO(FListe,i) ;
      MoveCurProgressForm(traduireMemoire('Interrogation du portail eWS pour le dossier ') + lTOB.GetString('dos_nodossier') );
      Application.ProcessMessages ;
      if FExpertScan.eWSActiverScan(lTOB.GetString('dos_nodossier')) then
       CSetParamScanEWS(lTOB.GetString('dos_nodossier'),'X')
        else
         CSetParamScanEWS(lTOB.GetString('dos_nodossier'),'-') ;

     end ; // if
  end // if
   else
    begin
     lTOB := GetO(FListe) ;
     MoveCurProgressForm(traduireMemoire('Interrogation du portail eWS pour le dossier ') + lTOB.GetString('dos_nodossier') );
     Application.ProcessMessages ;
     if FExpertScan.eWSActiverScan(lTOB.GetString('dos_nodossier')) then
      CSetParamScanEWS(lTOB.GetString('dos_nodossier'),'X')
       else
         CSetParamScanEWS(lTOB.GetString('dos_nodossier'),'-') ;
    end ;


 RefreshPclPge ;

 finally
  FiniMoveProgressForm ;
  SourisNormale ;
 end ;
end ;



procedure TOF_YYMULEXPERTSCAN.OnClickGenererArbo ( Sender : TObject) ;
var
 i    : integer ;
 lTob : TOB ;
 lInResult : integer ;

 procedure _Creer ;
 begin
  lInResult := FExpertScan.CreationArboDoss(lTOB.GetString('dos_nodossier')) ;
  if lInResult <> 0 then
   CSetParamScanHP(lTOB.GetString('dos_nodossier') , 'X' ) ;
 end ;

begin

 if PGIask('Confirmez-vous la création des répertoires pour Cegid Scan Expert ? ' , 'Création des répertoires') <> mrYes then exit ;

 SourisSablier ;
 InitMoveProgressForm(nil, '',traduireMemoire('Création des répertoires en cours...'), FListe.RowCount - 1 , False, True);

 lInResult := 0 ;

 try

 if ABoMultiSelected then
  begin
   for i := 1 to FListe.RowCount - 1 do
    if FListe.IsSelected(i) then
     begin
      lTob := GetO(FListe,i) ;
      _Creer ;
     end ; // if
  end // if
   else
    begin
     lTob := GetO(FListe) ;
     _Creer ;
    end ;

 if lInResult <> 0 then
  PGIInfo('Création des répertoires réussie') ;

 finally
  RefreshPclPge ;
  FiniMoveProgressForm ;
  SourisNormale ;
 end ;

end ;

procedure TOF_YYMULEXPERTSCAN.RemplitATobFListe ;
var
 lStWhere,lStWhere1,lStWhere2 : string ;
begin

 inherited ;

 if GetCheckBoxState('EWS_') = cbChecked then
  lStWhere  :=  ' and dos_ewscree="X"'
   else
    if GetCheckBoxState('EWS_') = cbUnChecked then
     lStWhere  :=  ' and dos_ewscree="-"' ;

 if GetCheckBoxState('SCANEWS_') = cbChecked then
  lStWhere1  :=  ' and yds_libelle="X"'
   else
    if GetCheckBoxState('SCANEWS_') = cbUnChecked then
     lStWhere1  :=  ' and ( (yds_libelle="-") or (yds_libelle is null) ) ' ; //' and yds_libelle="-"' ;

 if GetCheckBoxState('SCANHP_') = cbChecked then
  lStWhere2  :=  ' and yds_libre="X"'
   else
    if GetCheckBoxState('SCANHP_') = cbUnChecked then
     lStWhere2  :=  ' and ( ( yds_libre="-") or ( yds_libre is null) ) ' ;


 AStSqlTobFListe := 'select dos_nodossier,dos_libelle,dos_ewscree,yds_libelle BOSCANEWS,yds_libre BOHP from dossier ' +
                     'left outer join choixdpstd on ( yds_nodossier = dos_nodossier and yds_type="CES" ) ' +
                     ' where dos_cabinet<>"X" and dos_nodossier<>"000STD" '
                     + lStWhere + lStWhere1 + lStWhere2 ;


end;


procedure TOF_YYMULEXPERTSCAN.OnPopUpPopF11 (Sender: TObject);
var
 lPopUp  : TPopUpMenu ;
 i       : integer ;
 lTOB    : TOB ;
 lBoeWS  : boolean ;
begin

 if csDestroying in Ecran.ComponentState then Exit ;

 lPopUp := TPopUpMenu(GetControl('PopUpOutil', True)) ;
 lTob   := GetO(FListe,FListe.Row) ;
 if lTOB <> nil then
  lBoeWS := ( lTOB.GetString('BOSCANEWS') = 'X' )
   else
    lBoeWS := false ;

 for i := 0 to lPopUp.Items.Count - 1 do
  begin

   if lPopUp.Items[i].Name = 'GENERERXML' then
    begin
     lPopUp.Items[i].Enabled := not ABoMultiSelected ;
     Continue;
    end; // if

   if lPopUp.Items[i].Name = 'SYNCHROEWS' then
    begin
     lPopUp.Items[i].Enabled := FExpertScan.eWSActif and lBoeWS;
     Continue;
    end; // if

   if lPopUp.Items[i].Name = 'RECUPEREWS' then
    begin
     lPopUp.Items[i].Enabled := FExpertScan.eWSActif and lBoeWS ;
     Continue;
    end; // if

   if lPopUp.Items[i].Name = 'DEBLOQUEREWS' then
    begin
     lPopUp.Items[i].Enabled := FExpertScan.eWSActif and lBoeWS ;
     Continue;
    end; // if

   if lPopUp.Items[i].Name = 'ACTIVEREWS' then
    begin
     lPopUp.Items[i].Enabled := FExpertScan.eWSActif ;
     Continue;
    end; // if

  end ; // for

  inherited ;


end;

procedure TOF_YYMULEXPERTSCAN.OnErrorScan (sender : TObject; Error : TRecErrorScan ) ;
var
 lMess : string ;
begin
 Trace.TraceError('EXPERTSCAN',lMess) ;
 lMess := Error.RC_Message ;
 if V_PGI.SAV then
  lMess := lMess + Error.RC_Methode ;
 PGIError(lMess) ;
end ;

procedure TOF_YYMULEXPERTSCAN.OnProgress ( vStDossier,vStNature, vStMessage : string ) ;
begin
 Trace.TraceInformation('EXPERTSCAN',vStMessage);
 if vStMessage <> '' then
  MoveCurProgressForm(vStMessage)
   else
    MoveCurProgressForm(traduireMemoire('Récupération du dossier ') + vStDossier + ' nature : ' + vStNature );
end ;

procedure TOF_YYMULEXPERTSCAN.OnInfo ( vStMessage : string ) ;
begin
 PGIInfo(vStMessage) ;
end ;

procedure TOF_YYMULEXPERTSCAN.OnClickBCherche(Sender: TObject);
begin
 inherited;
 if not FExpertScan.eWSActif then
  begin
   FListe.ColWidths[3] := 0 ;
   FListe.ColWidths[4] := 0 ;
   FListe.ColWidths[5] := 0 ;
   FListe.ColWidths[6] := 0 ;
   THSystemMenu(GetControl('HMTrad')).ResizeGridColumns(FListe);
  end ;

end;

Initialization
  registerclasses ( [ TOF_YYMULEXPERTSCAN ] ) ;

end.
