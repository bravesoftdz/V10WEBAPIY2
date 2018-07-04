{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 07/02/2001
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : UTILISAT (UTILISAT)
Mots clefs ... : TOM;UTILISAT
*****************************************************************
PT1 | 25/03/2008 | V_803 | FL | Rattachement de US_AUXILIAIRE à la tablette PGSALARIEINT si gestion des intervenants extérieurs
}
Unit UTOMUTILISAT ;

Interface

Uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
     HCtrls, HEnt1, HMsgBox, UTOM, HTB97,math, paramsoc,
{$IFNDEF EAGLCLIENT}
     db,HDB,
    {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},
     Fiche, FichList, UserAcc,  Fe_Main,
{$ELSE}
     eFiche, eFichList,  Maineagl,
{$ENDIF}
{$IFDEF GCGC}
     EntGC,
{$ENDIF}
   variants,
  {$IFNDEF PGIMAJVER}
  {$IFDEF GCGC}
  {$IFDEF STK}
  StockUtil,
  {$ENDIF STK}
  {$ENDIF GCGC}
  {$ENDIF !PGIMAJVER}
     UTob, hsysmenu, hcrc, LicUtil, windows,ent1 ;

Type
  TOM_UTILISAT = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnArgument ( S: String )   ; override ;
    Procedure OnChangeField(F: TField)   ; override ;
    procedure OnClose                    ; override ;

  private
    tDroitAgenda          :TTreeView;
    iModeDroitAgenda      :integer;
    iLastCharged          :integer;
    bDroitAgendaModified  :boolean;
    TOBDroits             :TOB;
    TOBLibGroupes         :TOB; // $$$ JP 28/07/06
    sOnglet               :string;
    m_bSeriaAgenda        :boolean;
    m_bUserModifiable     :boolean;
{$IFDEF EWS}
    m_bEwsActif           :boolean;
{$ENDIF}
    bUtilisateurEws       :Boolean;
    bOnLoading            :Boolean;
    OldPwd : string;

    {$IFDEF BUREAU}
    bModifierDroitObligation : boolean;
    {$ENDIF BUREAU}

    {$IFDEF BUREAU}
    procedure AppelerClick (Sender:TObject);
    {$ENDIF}

    procedure ChargeMagUser;
    procedure PwdExit(Sender: TObject);
    procedure EPwdExit(Sender: TObject);
    procedure SPwdExit(Sender: TObject);

    procedure UpdateAgendaPanel;

    {$IFDEF BUREAU}
    procedure OnClickTypeObligation (Sender : TObject);
    {$ENDIF BUREAU}

    procedure OnClickUtilVersAutres (Sender:TObject);
    procedure OnClickAutresVersUtil (Sender:TObject);

    procedure OnDroitKeyDown        (Sender:TObject; var Key:Word; Shift:TShiftState);
    procedure OnDroitPopup          (Sender:TObject);
    procedure OnClickDroitAucun     (Sender:TObject);
    procedure OnClickDroitConsult   (Sender:TObject);
    procedure OnClickDroitAbsence   (Sender:TObject);
    procedure OnClickDroitActivite  (Sender:TObject);
    procedure OnClickDroitTout      (Sender:TObject);

    function  ComputeNewDroit       (iOldDroit:integer):integer;
    procedure SetDroitsAgendaMode   (iMode:integer);
    procedure DrawDroit             (DroitNode:TTreeNode);
    procedure DrawAllDroits         (TOBUnDroit:TOB=nil);
    procedure UpdateDroit           (TOBUnDroit:TOB; iNewDroit:integer=-1; bCanRefresh:boolean=FALSE);
    procedure Pages_OnChange        (Sender : TObject);
    procedure BImprimer_OnClick     (Sender : TObject);

    procedure UpdateDroitNode       (iNewDroit:integer=-1); //; bAlerteIfMulti:boolean=TRUE);
    
  public
    PW1, PW2, EPW1, EPW2, SPW1, SPW2 : THEdit ;
    sGroupeOrig :String;   // FQ N° 13632
    superviseur: boolean ;
    end ;

// $$$ JP 29/07/04 - il faut renvoyer le code retour du mul (pour sélection)
//Procedure FicheUSer(Quel : String) ;
function FicheUser (Quel:string; strArgument:string=''):string;

Implementation

uses
{$IFDEF DP}
        EntDP,
{$ENDIF}
{$IFDEF BUREAU}
 ObligationFonction,
{$ENDIF}
{$IFDEF EWS}
        UtileWS,
{$ENDIF}
        hstatus,
{$IFDEF GCGC}
        YRessource, // $$$ JP 23/04/07: pour gestion table unifiée user/ressource/salarié
{$ENDIF}
        menus;


// $$$ JP 29/07/04 - il faut renvoyer le code retour du mul (pour sélection)
{Procedure FicheUSer(Quel : String) ;
begin
//AGLLanceFiche('YY','YYUTILISATEURS','',Quel,'') ;
AGLLanceFiche('YY','YYUTILISAT_MUL','','','') ;
end;}
function FicheUser (Quel:string; strArgument:string=''):string;
begin
     Result := AGLLanceFiche ('YY','YYUTILISAT_MUL', '', '', strArgument);
end;

function ColleZeroDevant(Nombre , LongChaine : integer) : string ;
var
tabResult : string ;
TabInt : string;
i,j : integer;
begin
tabResult := '';
   for i := 1 to LongChaine do begin
      if Nombre < power(10,i) then
      begin
         TabInt := inttostr(Nombre);
        // colle (LongChaine-i zéro devant]
         for j := 0 to  (LongChaine-i-1)
                      do insert('0',TabResult,j);
         result := concat(TabResult,Tabint);
         exit;
      end;
    if i > LongChaine then result := inttostr(Nombre);
   end;
end;

Procedure TOM_UTILISAT.OnChangeField(F: TField)  ;
begin
inherited ;
If (F.FieldName='US_AUXILIAIRE') then
   Begin
   if (isnumeric(GetField('US_AUXILIAIRE')) AND (GetParamSoc('SO_PGTYPENUMSAL')='NUM')AND (Length(GetField('US_AUXILIAIRE'))<10)) then
        SetField('US_AUXILIAIRE',ColleZeroDevant(GetField('US_AUXILIAIRE'),10));
   end;
end;

procedure TOM_UTILISAT.OnClose;
begin
     inherited;

     if tDroitAgenda <> nil then
        tDroitAgenda.Items.Clear;

     TOBLibGroupes.Free; // $$$ JP 28/07/06
     TOBDroits.Free;
end;

procedure TOM_UTILISAT.OnNewRecord ;
begin
  inherited;
  // 18/05/2006 Lorsqu'on ouvre la fiche par le DispathTT, le champ est renseigné avec US_
  SetField('US_UTILISATEUR','');
  if VH^.CPIFDEFCEGID then
    SetField('US_GROUPE','INV');
    bUtilisateurEws := False;
end;

procedure TOM_UTILISAT.OnDeleteRecord;
var
  SQL                : String ;
  STC                : String ;
  QLoc               : TQuery ;
begin
  inherited;
  SQL := 'Select MO_TYPE,MO_NATURE,MO_CODE From MODELES Where MO_USER="' + GetField ( 'US_UTILISATEUR' ) + '" And MO_PERSO="X"' ;
  QLoc := OpenSql ( SQL, True ) ;
  While Not QLoc.Eof do
    begin
    StC:=QLoc.Fields[0].AsString+QLoc.Fields[1].AsString+QLoc.Fields[2].AsString ;
    ExecuteSql ( 'Delete From MODEDATA Where MD_CLE Like"' + StC + '%"' ) ;
    QLoc.Next ;
    end ;
  Ferme ( QLoc ) ;

  // $$$JP 20/01/04: ne pas faire de traitement sur les droits agenda si agenda non sérialisé
  if m_bSeriaAgenda = TRUE then
   begin
     // Il faut également supprimer les droits agenda de l'utilisateur, et les droits des autres sur l'utilisateur
     ExecuteSQL ('DELETE FROM CHOIXEXT WHERE YX_TYPE="DAU" AND (YX_LIBELLE="' + GetField ('US_UTILISATEUR') + '" OR YX_ABREGE="' + GetField ('US_UTILISATEUR') + '")');
     // et les droits sur la messagerie d'autres utilisateurs
     ExecuteSQL ('DELETE FROM YDATATYPETREES WHERE YDT_CODEHDTLINK="YYUSERMASTERSLAVE" AND YDT_MCODE="' + GetField ('US_UTILISATEUR') + '"');
   end;

   // $$$ JP 23/04/07: suppression dans la table unifiée
{$IFDEF GCGC}
   DeleteYRS ('', '', GetField ('US_UTILISATEUR'));
{$ENDIF}

{$IFDEF BUREAU}
  // MD 26/02/07 - Onglet "Obligations"
  ExecuteSQL ('DELETE FROM CHOIXCOD WHERE CC_TYPE="DPO" AND CC_CODE="' + GetField ('US_UTILISATEUR') + '"');
{$ENDIF}

  // $$$MD 13/01/05: gestion de l'espace eWS (ccmx)
  if bUtilisateurEws then
{$IFDEF EWS}
    EwsSupprimeCollab( GetField('US_UTILISATEUR'), GetField('US_ABREGE') );
{$ELSE}
    begin
    PGIInfo('Cet utilisateur a un espace eWS associé. Utilisez une application Cegid compatible eWS pour le supprimer.');
    LastError := -1 ;
    end;
{$ENDIF}

   // MD 07/03/07 - Suppression rattachement à des groupes de données
   // MB 18/04/2007
   ExecuteSQL('DELETE FROM LIENDONNEES WHERE LND_USERID="' + GetField ('US_UTILISATEUR') + '"');
{$IFDEF GRC}
   if (ctxGRC in V_PGI.PGIContexte) and
         GetParamsocSecur('SO_RTCONFIDENTIALITE', false) then
      ExecuteSQL('DELETE FROM PROSPECTCONF WHERE RTC_INTERVENANT="' + GetField ('US_UTILISATEUR') + '"');
{$ENDIF GRC}

end;

procedure TOM_UTILISAT.UpdateAgendaPanel;
begin
     // Libellés des radio-boutons adaptés aux noms utilisateurs, sauf si nouveau
     if DS.State <> dsInsert then
     begin
          SetControlCaption ('RBUTILVERSAUTRES', 'de ' + UpperCase (GetField ('US_LIBELLE')) + ' sur l''agenda des autres utilisateurs');
          SetControlCaption ('RBAUTRESVERSUTIL', 'des autres utilisateurs sur l''agenda de ' + UpperCase (GetField ('US_LIBELLE')));
     end
     else
     begin
          SetControlCaption ('RBUTILVERSAUTRES', 'du nouvel utilisateur sur l''agenda des autres utilisateurs');
          SetControlCaption ('RBAUTRESVERSUTIL', 'des autres utilisateurs sur l''agenda du nouvel utilisateur');
     end;

     // Si agenda sérialisé et modif autorisé, on active l'arborescence et on l'alimente par défaut de l'utilisateur vers les autres
     if (m_bSeriaAgenda = TRUE) and (m_bUserModifiable = TRUE) and (ds.State <> dsInsert) then
     begin
          SetControlEnabled ('GBVISUDROIT', TRUE);
          if TRadioButton (GetControl ('RBUTILVERSAUTRES')).Checked = FALSE then
              SetControlChecked ('RBUTILVERSAUTRES', TRUE) ;
          // MB : Chargement lors de l'arrivée sur l'onglet.
          // SetDroitsAgendaMode (0);
          SetControlEnabled ('LEXPLICDROITAGENDA', TRUE);
          SetControlEnabled ('LEXPLICDROITAGENDA2', TRUE);
     end
     else
     begin
          SetControlEnabled ('GBVISUDROIT', FALSE);
//          if (m_bSeriaAgenda = TRUE) and (ds.State <> dsInsert) then
//              SetDroitsAgendaMode (0)
//          else
//              SetDroitsAgendaMode (-1);
          SetControlEnabled ('LEXPLICDROITAGENDA', FALSE);
          SetControlEnabled ('LEXPLICDROITAGENDA2', FALSE);
     end;
end;

procedure TOM_UTILISAT.OnUpdateRecord ;
var Q                  : TQuery ;
    {$IFDEF BUREAU}
    TOBDroitObligation : Tob;
    ChChaine           : String;
    {$ENDIF BUREAU}
begin
  Inherited ;
  TFFiche(ecran).Retour := '';
  if GetControlText ( 'US_ABREGE' ) = '' then
    begin
    SetFocusControl ( 'US_ABREGE' ) ;
    HShowMessage ( '6;Utilisateurs;Vous devez renseigner un login.;W;O;O;O;', '', '' ) ;
    LastError := -1 ;
    Exit ;
    End ;

  if ExisteSQL('Select US_ABREGE from UTILISAT WHERE US_ABREGE="' + GetField ('US_ABREGE') +
                 '" AND US_UTILISATEUR<>"'+GetField('US_UTILISATEUR')+'"') then
  begin
    SetFocusControl ( 'US_ABREGE' ) ;
    HShowMessage ( '6;Utilisateurs;Le login existe déjà pour un autre utilisateur;W;O;O;O;', '', '' ) ;
    LastError := -1 ;
    Exit ;
  end;

  if GetControlText ( 'PW1' ) <> GetControlText ( 'PW2') then
    begin
    SetFocusControl ( 'PW2' ) ;
    HShowMessage ( '9;Utilisateurs;La confirmation est différente du mot de passe.;W;O;O;O;', '', '') ;
    LastError := -1 ;
    Exit ;
    End ;

  if GetControlText ( 'EPW1' ) <> GetControlText ( 'EPW2' ) then
    begin
    SetFocusControl ( 'EPW2' ) ;
    HShowMessage ( '9;Utilisateurs;La confirmation est différente du mot de passe.;W;O;O;O;', '', '') ;
    LastError := -1 ;
    Exit ;
    End ;

  if Assigned( GetControl ( 'SPW1' ) ) then
    begin
    if GetControlText ( 'SPW1' ) <> GetControlText ( 'SPW2') then
      begin
      SetFocusControl ( 'SPW2' ) ;
      HShowMessage ( '9;Utilisateurs;La confirmation est différente du mot de passe.;W;O;O;O;', '', '') ;
      LastError := -1 ;
      Exit ;
      End ;
    end;

  if GetControlText ( 'US_GROUPE' ) = '' then
    begin
    SetFocusControl ( 'US_GROUPE' ) ;
    HShowMessage ( '8;Utilisateurs;Vous devez renseigner un groupe d''utilisateurs.;W;O;O;O;', '', '') ;
    LastError := -1 ;
    Exit ;
    End ;

  SetControlText ( 'US_ABREGE', UpperCase ( GetControlText ( 'US_ABREGE' ) ) ) ;

  if GetControlText ( 'PW1' ) = '' then
  begin             //passwd à ''  => JSI le 03/06/04
{$IFDEF EWS}
    // MD 28/03/07 - ADM est utilisé par eWS 2.30 => mot de passe vide interdit
    if m_bEwsActif and ( GetControlText ('US_ABREGE')='ADM' ) then
    begin
         PGIInfo('Vous devez renseigner un mot de passe pour cet utilisateur.', Ecran.Caption);
         SetFocusControl('PW1');
         LastError := -1 ;
         exit;
    end;
{$ENDIF}
    Case PGIAskCancel('Pour des raisons de sécurité, il est fortement conseillé'
       + '#13 de renseigner le mot de passe.#13 Voulez-vous continuer ?', Ecran.Caption) of
           mrCancel : begin
                        LastError := -1 ;
                        exit;
                      end;
           mrNo : begin
                    SetFocusControl('PW1');
                    LastError := -1;
                    exit;
                  end;
           end;
  end;

  // $$$MD 13/01/05: gestion de l'espace eWS (ccmx)
  if bUtilisateurEws then
{$IFDEF EWS}
    EwsModifieCollab( GetField('US_ABREGE'), DecryptageSt(GetField('US_PASSWORD')), GetField('US_LIBELLE'), GetField('US_EMAIL') );
{$ELSE}
    begin
    PGIInfo('Cet utilisateur a un espace eWS associé. Utilisez une application PGI compatible eWS pour le modifier.');
    LastError := -1 ;
    exit;
    end;
{$ENDIF}

  if PW1.Text <> OldPwd then
    SetField('US_DATECHANGEPWD', Date);

  SetField ( 'US_PASSWORD', CryptageSt ( PW1.Text ) ) ;

  SetField ( 'US_EMAILPASSWORD', CryptageSt ( EPW1.Text ) ) ;

  if Assigned ( GetControl ('SPW1') ) then
    SetField ( 'US_EMAILSMTPPWD', CryptageSt ( SPW1.Text ) ) ;

  if ( GetControlText ( 'US_SUPERVISEUR' ) = '-' ) and
    Not ExisteSQL('select US_SUPERVISEUR from UTILISAT WHERE US_SUPERVISEUR="X" AND US_UTILISATEUR<>"'+GetField('US_UTILISATEUR')+'"') then
      SetControlText ( 'US_SUPERVISEUR', 'X' ) ;
  ChargeMagUser ;

  // Mise à jour table CHOIXEXT pour droits agenda
  if bDroitAgendaModified = TRUE then
  begin
    // Enregistrement des droits dans la base
    TOBDroits.InsertOrUpdateDB;
    bDroitAgendaModified := FALSE;
    SetControlEnabled ('RBUTILVERSAUTRES', TRUE);
    SetControlEnabled ('RBAUTRESVERSUTIL', TRUE);
  end;

  //JSI déplacé du OnAfterUpdateRecord le 16/06/04 pour pb CWAS
  Q := OpenSQL ( 'SELECT * FROM UTILISAT WHERE US_UTILISATEUR="' + GetField('US_UTILISATEUR')+ '"', True ) ;
  SetField ( 'US_CRC', Integer(GetCRC32ForData( Q )) ) ;
  Ferme ( Q ) ;

  {$IFDEF BUREAU}
  // Mise à jour table ChoixCod^pour droits obligation
  if (bModifierDroitObligation) then
   begin
    TOBDroitObligation := TOB.Create ('CHOIXCOD', nil, -1);
    TobDroitObligation.PutValue ('CC_TYPE','DPO');
    TobDroitObligation.PutValue ('CC_CODE',GetField('US_UTILISATEUR'));
    TobDroitObligation.ChargeCle1;
    TobDroitObligation.LoadDB;
    ChChaine:='';
    if (TCheckBox (GetControl ('CHB_FISCALE')).Checked) then ChChaine:=ChChaine+'FIS;';
    if (TCheckBox (GetControl ('CHB_SOCIALE')).Checked) then ChChaine:=ChChaine+'SOC;';
    if (TCheckBox (GetControl ('CHB_JURIDIQUE')).Checked) then ChChaine:=ChChaine+'JUR';
    TobDroitObligation.PutValue ('CC_LIBELLE',ChChaine);
    TobDroitObligation.InsertOrUpdateDB;
    TobDroitObligation.Free;
   end;
  {$ENDIF}

  TFFiche(ecran).Retour := GetField('US_UTILISATEUR');
end ;

procedure TOM_UTILISAT.OnAfterUpdateRecord;
var
{$IFDEF EWS}
   TOBUser  :TOB;
{$ENDIF}
   strUser  :string;
begin
     inherited;

     // $$$JP 13/07/04: mise à jour panel Droits Agenda
     UpdateAgendaPanel;

     // $$$ JP 23/04/07
     strUser := GetField ('US_UTILISATEUR');

    // Début FQ N° 13632
    if (sGroupeOrig <> GetField('US_GROUPE')) And (ExisteSQL('select CPU_USER from CPPROFILUSERC WHERE CPU_TYPE="ETA" AND CPU_USER="'+GetField('US_UTILISATEUR')+'" AND CPU_USERGRP="'+ sGroupeOrig +'"')) then
       ExecuteSQL('UPDATE CPPROFILUSERC SET CPU_USERGRP="' + GetField('US_GROUPE') + '" WHERE CPU_TYPE="ETA" AND CPU_USER="'+strUser+'" AND CPU_USERGRP="'+ sGroupeOrig +'"'); // $$$ JP 23/04/07: strUser
    // Fin FQ N° 13632

    // $$$ JP 23/04/07: synchronisation de la table YRESSOURCE (ajout ou modif')
{$IFDEF GCGC}
    CreateYRS ('', '', strUser);
{$ENDIF}

{$IFDEF EWS}
   // $$$ JP 14/12/06: si le groupe est typé "ews", on créer immédiatement l'utilisateur dans Ews
   if (m_bEwsActif = TRUE) and  (ExisteSQL ('SELECT 1 FROM ##DP##.CHOIXCOD WHERE CC_TYPE="GEW" AND CC_CODE="' + GetField ('US_GROUPE') + '"') = TRUE) then
   begin
        TOBUser := TOB.Create ('le user', nil, -1);
        try
           TOBUser.AddChampSupValeur ('US_UTILISATEUR', GetField ('US_UTILISATEUR'));
           TOBUser.AddChampSupValeur ('US_ABREGE',      GetField ('US_ABREGE'));
           TOBUser.AddChampSupValeur ('US_PASSWORD',    GetField ('US_PASSWORD'));
           TOBUser.AddChampSupValeur ('US_LIBELLE',     GetField ('US_LIBELLE'));
           TOBUser.AddChampSupValeur ('US_EMAIL',       GetField ('US_EMAIL'));
           EwsVerifOuCreeCollab (TOBUser);
        finally
          TOBUser.Free;
        end;
   end;
{$ENDIF}
end ;

procedure TOM_UTILISAT.OnLoadRecord ;
var
  connecte: boolean ;
begin
  Inherited ;
  Connecte := GetField ( 'US_PRESENT' ) = 'X' ;
  SetControlEnabled('BDelete',not Connecte) ;
  if Not Superviseur then FicheReadOnly ( ECRAN ) ;
  SetControlVisible('TCONNECTED',Connecte) ;
  SetControlVisible('TNOTCONNECTED',Not Connecte) ;
  SetControlVisible('US_DATECONNEXION',Connecte) ;
  PW1.Text := DeCryptageSt ( GetField ( 'US_PASSWORD' ) ) ;
  PW2.Text := DeCryptageSt ( GetField ( 'US_PASSWORD' ) ) ;
  EPW1.Text := DeCryptageSt ( GetField ( 'US_EMAILPASSWORD' ) ) ;
  EPW2.Text := DeCryptageSt ( GetField ( 'US_EMAILPASSWORD' ) ) ;
  OldPwd := PW1.Text;
  if Assigned ( GetControl ( 'SPW1' ) ) then
    begin
    SPW1.Text := DeCryptageSt ( GetField ( 'US_EMAILSMTPPWD' ) ) ;
    SPW2.Text := DeCryptageSt ( GetField ( 'US_EMAILSMTPPWD' ) ) ;
    end;

  // Libellé des options d'affichage des droits agenda
  // $$$JP 20/01/04: ne pas faire de traitement sur les droits agenda si agenda non sérialisé
  // $$$JP 13/07/04: ne rien faire si en mode insertion (+ transfert dans une fonction)
  UpdateAgendaPanel;

  // Accès direct à un onglet
  if not (DS.State in [dsInsert]) then
  begin
       if (sOnglet='PINTERNET') or (sOnglet='PDROITAGENDA') then
       begin
            TPageControl (GetControl ('PAGES')).ActivePage := TTabSheet (GetControl (sOnglet));
            Pages_OnChange (nil);
       end;
  end;

  // $$$MD 13/01/05: gestion de l'espace eWS (ccmx)
  bUtilisateurEws := ( GetField('US_UTILISATEUR')<>'' )
   and ExisteSQL('SELECT 1 FROM ##DP##.PARAMSOC WHERE SOC_NOM="SO_NE_EWSACTIF" AND SOC_DATA="X"') // eWS actif dans l'environnement
   and ExisteSQL('SELECT 1 FROM ##DP##.CHOIXCOD WHERE CC_TYPE="YEW" AND CC_CODE="'+GetField('US_UTILISATEUR')+'"'); // et utilisateur été activé sur eWS


  //--- CAT : 04/10/2005 Gestion des droits obligations
  {$IFDEF BUREAU}
   if (DonnerAccesObligation ('FIS',GetField ('US_UTILISATEUR'))) then TCheckBox (GetControl ('CHB_FISCALE')).Checked:=True;
   if (DonnerAccesObligation ('SOC',GetField ('US_UTILISATEUR'))) then TCheckBox (GetControl ('CHB_SOCIALE')).Checked:=True;
   if (DonnerAccesObligation ('JUR',GetField ('US_UTILISATEUR'))) then TCheckBox (GetControl ('CHB_JURIDIQUE')).Checked:=True;
  {$ENDIF}

  sGroupeOrig := GetField('US_GROUPE'); // FQ N° 13632
  //uniquement en line
  {*
    SetControlVisible('US_FONCTION', false);
    SetControlVisible('TUS_FONCTION', false);
    SetControlVisible('US_GROUPE', false);
    SetControlVisible('TUS_GROUPE', false);
    SetControlVisible('US_AUXILIAIRE', false);
    SetControlVisible('TARS_SALARIE', false);
    SetControlVisible('US_SUPERVISEUR', false);
    SetControlVisible('US_CONTROLEUR', false);
    SetControlVisible('US_GRPSDELEGUES', false);
    SetControlVisible('TUS_GRPSDELEGUES', false);
  *}

  bOnLoading := False;
end;

procedure TOM_UTILISAT.PwdExit ( Sender: TObject ) ;
begin
if CryptageSt ( PW1.Text ) <> GetField ( 'US_PASSWORD' ) then
   begin
   if not(DS.State in [dsInsert,dsEdit])then DS.edit;
   setfield ( 'US_PASSWORD', CryptageSt ( PW1.Text ) ) ;
   end;
end ;

procedure TOM_UTILISAT.EPwdExit ( Sender: TObject ) ;
begin
if CryptageSt ( EPW1.Text ) <> GetField ( 'US_EMAILPASSWORD' ) then
   begin
   if not(DS.State in [dsInsert,dsEdit])then DS.edit;
   SetField ( 'US_EMAILPASSWORD', CryptageSt ( EPW1.Text ) ) ;
   end;
end ;

procedure TOM_UTILISAT.SPwdExit ( Sender: TObject ) ;
begin
if CryptageSt ( SPW1.Text ) <> GetField ( 'US_EMAILSMTPPWD' ) then
   begin
   if not(DS.State in [dsInsert,dsEdit])then DS.edit;
   SetField ( 'US_EMAILSMTPPWD', CryptageSt ( SPW1.Text ) ) ;
   end;
end ;

function TOM_UTILISAT.ComputeNewDroit (iOldDroit:integer):integer;
begin
     // Calcul du nouveau droit
     Result := 0;
     if (iOldDroit >= 0) and (iOldDroit <= 4) then
        if iOldDroit = 4 then
            Result := 0
        else
            Result := iOldDroit + 1;
end;

procedure TOM_UTILISAT.UpdateDroitNode (iNewDroit:integer=-1); //; bAlerteIfMulti:boolean=TRUE);
var
   CurNode      :TTreeNode;
   FilsNode     :TTreeNode;
   TOBUnDroit   :TOB;
   strDroit     :string;
   iDroit       :integer;
   bCanUpdate   :boolean;
begin
  iDroit := 0; // DBR mais franchement a quoi sert iDroit ?
     // Si la fiche est en consultation seulement, on ne fait pas de màj des droits agenda
     if m_bUserModifiable = FALSE then
        exit;

     CurNode := tDroitAgenda.Selected;
     if CurNode <> nil then
     begin
          tDroitAgenda.Items.BeginUpdate;

          // Mise à jour du droit dans l'arborescence
          TOBUnDroit := TOB (CurNode.Data);
          if TOBUnDroit <> nil then
          begin
               if iNewDroit = -1 then
                  iNewDroit := ComputeNewDroit (StrToInt (TOBUnDroit.GetValue ('YX_LIBRE')));
               UpdateDroit (TOBUnDroit, iNewDroit, TRUE);
          end
          else
          begin
               if (CurNode <> tDroitAgenda.Items.GetFirstNode) then
               begin
                    FilsNode := CurNode.GetFirstChild;
                    if FilsNode <> nil then
                    begin
                         TOBUnDroit := TOB (FilsNode.Data);
                         if iNewDroit = -1 then
                            iNewDroit := ComputeNewDroit (StrToInt (TOBUnDroit.GetValue ('YX_LIBRE')));

                         // Si alerte utilisateur demandée, on construit le message et on demande confirmation
                         bCanUpdate := TRUE;
//                         if balerteIfMulti = TRUE then
  //                       begin
                              // construction message d'avertissement du changement des droits pour le groupe entier
                              if iModeDroitAgenda = 0 then
                                  strDroit := 'L''agenda des utilisateurs du groupe "' + CurNode.Text + '" vont devenir '
                              else
                                  strDroit := 'L''agenda de ' + GetField ('US_LIBELLE') + ' va devenir ';
                              case iNewDroit of
                                   0:
                                     strDroit := strDroit + 'INVISIBLE ';
                                   1:
                                     strDroit := strDroit + 'CONSULTABLE ';
                                   2:
                                     strDroit := strDroit + 'MODIFIABLE POUR ABSENCE ';
                                   3:
                                     strDroit := strDroit + 'MODIFIABLE POUR TRAVAIL ';
                                   4:
                                     strDroit := strDroit + 'TOTALEMENT MODIFIABLE ';
                              end;
                              if iModeDroitAgenda = 0 then
                              begin
                                   if iDroit = 0 then
                                       strDroit := strDroit + 'pour '
                                   else
                                       strDroit := strDroit + 'par ';
                                   strDroit := strDroit + GetField ('US_LIBELLE');
                              end
                              else
                              begin
                                   if iDroit = 0 then
                                       strDroit := strDroit + 'pour les '
                                   else
                                       strDroit := strDroit + 'par les ';
                                   strDroit := strDroit + 'utilisateurs du groupe "' + CurNode.Text + '"';
                              end;

                              // Si confirmation, mise à jour des droits des utilisateurs "enfants" dans l'arborescence
                              if (PgiAsk (strDroit + #10 + ' Confirmez-vous cette opération?') <> mrYes) then
                                 bCanUpdate := FALSE;
    //                     end;

                         // Si màj autorisée, on lance sur tous les fils du noeud
                         if bCanUpdate = TRUE then
                         begin
                              CurNode := FilsNode;
                              while CurNode <> nil do
                              begin
                                   TOBUnDroit := TOB (CurNode.Data);
                                   if TOBUnDroit <> nil then
                                      UpdateDroit (TOBUnDroit, iNewDroit);
                                   CurNode := CurNode.GetNextSibling
                              end;

                              // màj de tous les droits de l'arborescence
                              DrawAllDroits;
                         end;
                    end;
               end;
          end;

          tDroitAgenda.Selected.Expand (TRUE);
          tDroitAgenda.Items.EndUpdate;
     end;

end;

{$IFDEF BUREAU}
procedure TOM_UTILISAT.OnClickTypeObligation (Sender : TObject);
begin
 if bOnLoading then exit;
 bModifierDroitObligation:=True;
 if not (DS.State in [dsInsert, dsEdit]) then DS.Edit;
 // MD 26/02/07 - Obligation de faire une modif factice, mais depuis quand ?
 SetField('US_LIBELLE', GetControlText('US_LIBELLE'));
end;
{$ENDIF BUREAU}

procedure TOM_UTILISAT.OnClickUtilVersAutres(Sender: TObject);
begin
  if TPageControl(GetControl('PAGES')).ActivePageIndex = 3 then
     SetDroitsAgendaMode (0);
end;

procedure TOM_UTILISAT.OnClickAutresVersUtil(Sender: TObject);
begin
  if TPageControl(GetControl('PAGES')).ActivePageIndex = 3 then
     SetDroitsAgendaMode (1);
end;

procedure TOM_UTILISAT.OnDroitKeyDown (Sender:TObject; var Key:Word; Shift:TShiftState);
begin
     if Key = 32 then
        UpdateDroitNode;
end;

procedure TOM_UTILISAT.OnDroitPopup (Sender:TObject);
var
   CurNode     :TTreeNode;
   TOBUnDroit  :TOB;
   iDroit      :integer;
   bCanPopup    :boolean;
begin
     CurNode := tDroitAgenda.Selected;
     if CurNode <> nil then
     begin
          // On sélectionne définitivement l'élément (sinon, bizarrement retour sur le dernier sélectionné par click gauche)
          tDroitAgenda.Selected := CurNode;

          // Détermination du droit sur l'élément
          bCanPopup  := m_bUserModifiable;
          TOBUnDroit := TOB (CurNode.Data);
          if TOBUnDroit <> nil then
              iDroit := StrToInt (TOBUnDroit.GetValue ('YX_LIBRE'))
          else
              iDroit := -1;
     end
     else
     begin
          bCanPopup := FALSE;
          iDroit   := -1;
     end;

     // Màj apparence des menus du popup
     TMenuItem (GetControl ('ATD_AUCUN')).Enabled     := bCanPopup;
     TMenuItem (GetControl ('ATD_CONSULT')).Enabled   := bCanPopup;
     TMenuItem (GetControl ('ATD_ABSENCE')).Enabled   := bCanPopup;
     TMenuItem (GetControl ('ATD_ACTIVITE')).Enabled  := bCanPopup;
     TMenuItem (GetControl ('ATD_TOUT')).Enabled      := bCanPopup;

     // On coche le bon menu (droit de l'élément en sélection, s'il existe)
     TMenuItem (GetControl ('ATD_AUCUN')).Checked     := (iDroit = 0);
     TMenuItem (GetControl ('ATD_CONSULT')).Checked   := (iDroit = 1);
     TMenuItem (GetControl ('ATD_ABSENCE')).Checked   := (iDroit = 2);
     TMenuItem (GetControl ('ATD_ACTIVITE')).Checked  := (iDroit = 3);
     TMenuItem (GetControl ('ATD_TOUT')).Checked      := (iDroit = 4);
end;

procedure TOM_UTILISAT.OnClickDroitAucun (Sender:TObject);
begin
     UpdateDroitNode (0);
end;

procedure TOM_UTILISAT.OnClickDroitConsult (Sender:TObject);
begin
     UpdateDroitNode (1);
end;

procedure TOM_UTILISAT.OnClickDroitAbsence (Sender:TObject);
begin
     UpdateDroitNode (2);
end;

procedure TOM_UTILISAT.OnClickDroitActivite (Sender:TObject);
begin
     UpdateDroitNode (3);
end;

procedure TOM_UTILISAT.OnClickDroitTout (Sender:TObject);
begin
     UpdateDroitNode (4);
end;

procedure TOM_UTILISAT.SetDroitsAgendaMode (iMode:integer);
var
   TOBUser        :TOB;
   TOBUnDroit     :TOB;
   TOBUnGroupe    :TOB;
   Q              :TQuery;
   i              :integer;
   RacineNode     :TTreeNode;
   GroupeNode     :TTreeNode;
   UserNode       :TTreeNode;
   varGroupe      :string;
   varCurGroupe   :string;
   strMaster      :string;
   strSlave       :string;
begin
     if iLastCharged = iMode then exit ;
     // Raz des droits et arborescence, qui vont être reconstruits
     if tDroitAgenda <> nil then
        tDroitAgenda.Items.Clear;
     if TOBDroits <> nil then
        TOBDroits.ClearDetail;

     // Si non superviseur, pas d'affichage des droits agenda
     if not Superviseur then
        exit;

     // Si mode d'affichage non assigné, on fait rien
     iModeDroitAgenda := iMode;
     if iModeDroitAgenda = -1 then
        exit;

     // Lecture des droits agenda en mémoire pour l'utilisateur
     Q := nil;
     try
      if iModeDroitAgenda = 0 then
          Q := OpenSQL('SELECT YX_TYPE,YX_CODE,YX_LIBELLE,YX_ABREGE,YX_LIBRE FROM CHOIXEXT WHERE YX_TYPE="DAU" AND YX_LIBELLE="' + GetField ('US_UTILISATEUR') + '"', FALSE)
      else
          Q := OpenSQL ('SELECT YX_TYPE,YX_CODE,YX_LIBELLE,YX_ABREGE,YX_LIBRE FROM CHOIXEXT WHERE YX_TYPE="DAU" AND YX_ABREGE="' + GetField ('US_UTILISATEUR') + '"', FALSE);
      if Q <> nil then
         TOBDroits.LoadDetailDB('CHOIXEXT', '', '', Q, TRUE);
     finally
          Ferme (Q);
     end;

     // Définition des droits agenda
     tDroitAgenda.Items.BeginUpdate;
     if iModeDroitAgenda = 0 then
         RacineNode := tDroitAgenda.Items.Add (nil, 'Droits de ' + UpperCase (GetField ('US_LIBELLE')) + ' sur l''agenda des autres utilisateurs')
     else
         RacineNode := tDroitAgenda.Items.Add (nil, 'Droits des autres utilisateurs sur l''agenda de ' + UpperCase (GetField ('US_LIBELLE')));
     RacineNode.ImageIndex    := 94;
     RacineNode.SelectedIndex := 94;
     if RacineNode <> nil then
     begin
        TOBUser := TOB.Create ('les users', nil, -1);
        try
           TOBUser.LoadDetailFromSQL ('SELECT US_UTILISATEUR, US_LIBELLE, UCO_GROUPECONF FROM UTILISAT LEFT JOIN ##DP##.USERCONF ON US_UTILISATEUR=UCO_USER ORDER BY UCO_GROUPECONF, US_UTILISATEUR');
           VarCurGroupe := '';
           GroupeNode   := nil;
           InitMove (TOBUser.Detail.Count, 'Chargement des droits agenda');
           for i := 0 to TOBUser.Detail.Count - 1 do
           begin
                // Progression chargement droits
                MoveCur (FALSE);

                // Si rupture sur groupe (ou bien 1er groupe identifié), on insère ce nouveau groupe
                varGroupe := TOBUser.Detail [i].GetString ('UCO_GROUPECONF');

                // $$$ JP 25/04/06 - il faut avoir une clé unique, sur 6 caractères: 3 car. user maitre + 3 car. user esclave
                if iModeDroitAgenda = 0 then
                begin
                     // $$$ JP 16/08/06: nouvelle clé YX_CODE (avec séparateur #)
                     strMaster := Trim (GetField ('US_UTILISATEUR')); // Copy (GetField ('US_UTILISATEUR') + '   ', 1, 3);
                     strSlave  := Trim (TOBUser.Detail [i].GetValue ('US_UTILISATEUR')); // Copy (TOBUser.Detail [i].GetValue ('US_UTILISATEUR') + '   ', 1, 3);
                end
                else
                begin
                     // $$$ JP 16/08/06: nouvelle clé YX_CODE (avec séparateur #)
                     strMaster := Trim (TOBUser.Detail [i].GetValue ('US_UTILISATEUR')); //Copy (TOBUser.Detail [i].GetValue ('US_UTILISATEUR') + '   ', 1, 3);
                     strSlave  := Trim (GetField ('US_UTILISATEUR')); //Copy (GetField ('US_UTILISATEUR') + '   ', 1, 3);
                end;

                // $$$ JP 29/03/2005 - changement suite gestion interdite des null par agl 580
                if (i = 0) or (varGroupe <> varCurGroupe) then // if (varCurGroupe = '') or (varGroupe <> varCurGroupe) then
                begin
                     varCurGroupe := varGroupe;
                     if (varGroupe = '') then
                         GroupeNode := tDroitAgenda.Items.AddChild (RacineNode, 'Hors groupe de travail')
                     else
                     begin
                          // $$$ JP 28/07/06: pour éviter RechDom, on pioche désormais dans la TOB des libellés groupe de travail
                          TOBUnGroupe := TOBLibGroupes.FindFirst (['GRP_CODE'], [VarToStr (varGroupe)], TRUE);
                          if TOBUnGroupe <> nil then
                              GroupeNode := tDroitAgenda.Items.AddChild (RacineNode, Trim (TOBUnGroupe.GetString ('GRP_LIBELLE')))
                          else
                              GroupeNode := tDroitAgenda.Items.AddChild (RacineNode, VarToStr (varGroupe) + ': groupe de travail inconnu');
                     end;
                              //GroupeNode := tDroitAgenda.Items.AddChild (RacineNode, RechDom('YYGROUPECONF', UpperCase (varGroupe), FALSE));
                     GroupeNode.ImageIndex    := 95;
                     GroupeNode.SelectedIndex := 95;
                end;

                // Insertion utilisateur dans le groupe en cours
                UserNode := tDroitAgenda.Items.AddChild (GroupeNode, TOBUser.Detail [i].GetValue ('US_UTILISATEUR') + ' - ' + TOBUser.Detail [i].GetValue ('US_LIBELLE'));

                // Définition des droits inter-utilisateur (selon le mode d'affichage)
                if iModeDroitAgenda = 0 then
                    TOBUnDroit := TOBDroits.FindFirst (['YX_ABREGE'], [strSlave], TRUE) // $$$ JP 25/04/06 [TOBUser.Detail [i].GetValue ('US_UTILISATEUR')], TRUE)
                else
                    TOBUnDroit := TOBDroits.FindFirst (['YX_LIBELLE'], [strMaster], TRUE); // $$$ JP 25/04/06 [TOBUser.Detail [i].GetValue ('US_UTILISATEUR')], TRUE);
                if TOBUnDroit = nil then
                begin
                     UserNode.ImageIndex := 2-1;
                     TOBUnDroit := TOB.Create ('CHOIXEXT', TOBDroits, -1);
                     TOBUnDroit.AddChampSupValeur ('YX_TYPE', 'DAU');
                     {begin
                          TOBUnDroit.AddChampSupValeur ('YX_CODE', GetField ('US_UTILISATEUR')+TOBUser.Detail [i].GetValue ('US_UTILISATEUR'));
                          TOBUnDroit.AddChampSupValeur ('YX_LIBELLE', GetField ('US_UTILISATEUR'));
                          TOBUnDroit.AddChampSupValeur ('YX_ABREGE', TOBUser.Detail [i].GetValue ('US_UTILISATEUR'));
                     end
                     else
                     begin
                          TOBUnDroit.AddChampSupValeur ('YX_CODE', TOBUser.Detail [i].GetValue ('US_UTILISATEUR')+GetField ('US_UTILISATEUR'));
                          TOBUnDroit.AddChampSupValeur ('YX_LIBELLE', TOBUser.Detail [i].GetValue ('US_UTILISATEUR'));
                          TOBUnDroit.AddChampSupValeur ('YX_ABREGE', GetField ('US_UTILISATEUR'));
                     end;}
                     TOBUnDroit.AddChampSupValeur ('YX_CODE',    strMaster+'#'+strSlave);
                     TOBUnDroit.AddChampSupValeur ('YX_LIBELLE', strMaster);
                     TOBUnDroit.AddChampSupValeur ('YX_ABREGE',  strSlave);
                     TOBUnDroit.AddChampSupValeur ('YX_LIBRE',   '0');
                end;
                UserNode.Data := TOBUnDroit;

                // Màj des icones du droit nouvellement créé
                DrawDroit (UserNode);
           end;
           FiniMove;
        finally
               TOBUser.Free;
        end;
        RacineNode.Expand (TRUE);
        RacineNode.MakeVisible;
     end;
     tDroitAgenda.Items.EndUpdate;
     iLastCharged := iMode ;
end;

// Affichage des droits pour les noeuds référençant le droit spécifié (si nil, recharge tous les droits)
procedure TOM_UTILISAT.DrawAllDroits (TOBUnDroit:TOB);
var
   curNode   :TTreeNode;
begin
     // On passe en revue tous les noeuds, pour màj ceux correspondant au droit spécifié (un utilisateur pouvant appartenir à plusieurs groupe de travail)
     curNode := tDroitAgenda.Items.GetFirstNode;
     while curNode <> nil do
     begin
          if (TOBUnDroit = nil) or (TOB (curNode.Data) = TOBUnDroit) then
             DrawDroit (curNode);

          // Noeuds suivant
          CurNode := CurNode.GetNext;
     end;
end;

procedure TOM_UTILISAT.DrawDroit (DroitNode:TTreeNode);
begin
     // Détermination de l'icone correspondant au droit du noeud spécifié
     with DroitNode do
     begin
          if Data <> nil then
          begin
               case StrToInt (TOB (Data).GetString ('YX_LIBRE')) of
                    0:
                      ImageIndex := 1;
                    1:
                      ImageIndex := 93;
                    2:
                      ImageIndex := 32;
                    3:
                      ImageIndex := 28;
                    4:
                      ImageIndex := 11;
               end;
               SelectedIndex := ImageIndex;
          end;
     end;
end;

procedure TOM_UTILISAT.UpdateDroit (TOBUnDroit:TOB; iNewDroit:integer; bCanRefresh:boolean);
begin
     SetControlEnabled ('RBUTILVERSAUTRES', FALSE);
     SetControlEnabled ('RBAUTRESVERSUTIL', FALSE);

     // Màj du droit dans la TOB, puis màj de l'arborescence complète (car l'utilisateur peut apparaitre dans plusieurs groupe de travail)
     TOBUnDroit.PutValue ('YX_LIBRE', IntToStr (iNewDroit));

     // $$$ JP 29/03/2005: possibilité de ne pas rafraichir, dans le cas d'un boucle par exemple
     if bCanRefresh = TRUE then
        DrawAllDroits (TOBUnDroit);

     // Indique au moteur de donnée que la fiche a été modifiée
     bDroitAgendaModified := TRUE;
     if not (DS.State in [dsInsert, dsEdit]) then
        DS.Edit;
end;

procedure TOM_UTILISAT.OnArgument ( S: String ) ;
var Edit : {$IFNDEF EAGLCLIENT}THDBEdit{$ELSE}THEdit{$ENDIF}; //PT1
begin
  Inherited ; // traite action= mais ne l'enlève pas
  bOnLoading     := True;
  iLastCharged := -1 ;

  PW1            := THEdit ( GetControl ( 'PW1' ) ) ;
  PW2            := THEdit ( GetControl ( 'PW2' ) ) ;
  EPW1           := THEdit ( GetControl ( 'EPW1' ) ) ;
  EPW2           := THEdit ( GetControl ( 'EPW2' ) ) ;
  PW1.OnExit     := PwdExit ;
  EPW1.OnExit    := EPwdExit ;
  if Assigned(GetControl('SPW1')) then
    begin
    SPW1         := THEdit ( GetControl ( 'SPW1' ) ) ;
    SPW2         := THEdit ( GetControl ( 'SPW2' ) ) ;
    SPW1.OnExit  := SPwdExit ;
    end;
  Superviseur := (V_PGI.Superviseur = True ) ;
  if Not Superviseur then FicheReadOnly ( ECRAN ) ;
  SetControlVisible('US_SUIVILOG',V_PGI.LaSerie >= S5) ;


{$IFDEF BUREAU}
  bModifierDroitObligation:=False;
  TCheckBox (GetControl ('CHB_FISCALE')).OnClick:=OnClickTypeObligation;
  TCheckBox (GetControl ('CHB_SOCIALE')).OnClick:=OnClickTypeObligation;
  TCheckBox (GetControl ('CHB_JURIDIQUE')).OnClick:=OnClickTypeObligation;
{$ENDIF}

  // $$$JP 20/01/04: ne pas faire de traitement sur les droits agenda si agenda non sérialisé
{$IFDEF DP}
  m_bSeriaAgenda := VH_DP.SeriaMessagerie;
{$ELSE}
  m_bSeriaAgenda := FALSE;
  {$IFNDEF PGIMAJVER}
  {$IFDEF GCGC}
  {$IFDEF STK}
  if StkGereContremarque then
  begin
    SetControlVisible('TUS_EMAILLOGIN',False);
    SetControlVisible('US_EMAILLOGIN',False);
    SetControlVisible('TUS_EMAILPASSWORD',False);
    SetControlVisible('EPW1',False);
    SetControlVisible('TUS_PASSWORD2',False);
    SetControlVisible('EPW2',False);
    SetControlVisible('TUS_EMAILPOPSERVER',False);
    SetControlVisible('US_EMAILPOPSERVER',False);
    SetControlVisible('TUS_EMAILSMTPSERVER',False);
    SetControlVisible('US_EMAILSMTPSERVER',False);
  end else
  {$ENDIF STK}
  {$ENDIF GCGC}
  {$ENDIF !PGIMAJVER}
  SetControlVisible('PDROITAGENDA',False);
  SetControlVisible('PDROITOBLIGATION',False);
{$ENDIF}
{$IFDEF V10}
  SetControlVisible('PINTERNET', True);
{$ELSE}
  SetControlVisible('PINTERNET',False);
{$ENDIF}
  if m_bSeriaAgenda = TRUE then
  begin
       // On a besoin de la liste d'image standard PGI (pour les noeuds du treeview)
       if V_PGI.GraphList = nil then
          ChargeImageList;

       // Initialisation treeview des droits agenda
       tDroitAgenda := TTreeView (GetControl ('TVDROITAGENDA'));
       if tDroitAgenda <> NIL then
       begin
            tDroitAgenda.ShowButtons              := TRUE;
            tDroitAgenda.OnKeyDown                := OnDroitKeyDown;
            tDroitAgenda.Images                   := V_PGI.GraphList;
            tDroitAgenda.ReadOnly                 := TRUE;
            tDroitAgenda.RightClickSelect         := TRUE;
            tDroitAgenda.PopupMenu.OnPopup        := OnDroitPopup;

            TMenuItem (GetControl ('ATD_AUCUN')).Onclick    := OnClickDroitAucun;
            TMenuItem (GetControl ('ATD_CONSULT')).Onclick  := OnClickDroitConsult;
            TMenuItem (GetControl ('ATD_ABSENCE')).Onclick  := OnClickDroitAbsence;
            TMenuItem (GetControl ('ATD_ACTIVITE')).Onclick := OnClickDroitActivite;
            TMenuItem (GetControl ('ATD_TOUT')).Onclick     := OnClickDroitTout;
       end;

       // TOB des droits
       TOBDroits            := TOB.Create ('CHOIXEXT', nil, -1);
       bDroitAgendaModified := FALSE;

       // $$$ JP 28/07/06: liste des libellés groupe de travail (pour éviter RechDom dans une boucle)
       TOBLibGroupes        := TOB.Create ('les groupes', nil, -1);
       TOBLibGroupes.LoadDetailFromSQL ('SELECT GRP_CODE,GRP_LIBELLE FROM GRPDONNEES WHERE GRP_NOM="GROUPECONF" ORDER BY GRP_CODE');

       // Mode de visualisation des droits agenda
       TRadioButton (GetControl ('RBUTILVERSAUTRES')).OnClick := OnClickUtilVersAutres;
       TRadioButton (GetControl ('RBAUTRESVERSUTIL')).OnClick := OnClickAutresVersUtil;

       // Message dans le treeview si pas le droit de voir les droits
       if not Superviseur then
          tDroitAgenda.Items.Add (nil, 'Droits sur agenda non disponibles pour un utilisateur "non administrateur"');
  end
  else
  begin
       TOBDroits            := nil;
       TOBLibGroupes        := nil; // $$$ JP 28/07/06
       bDroitAgendaModified := FALSE;
  end;

  // $$$ JP - modif autorisée si pas ACTION=CONSULTATION
  m_bUserModifiable := Pos ('ACTION=CONSULTATION', UpperCase (S)) = 0;

  sOnglet := ReadTokenSt(S);
  if Copy(sOnglet, 1, 7)='ACTION=' then sOnglet := ReadTokenSt(S);

  TPageControl(GetControl('PAGES')).OnChange := Pages_OnChange;
  TToolBarButton97(GetControl('BImprimer')).OnClick := BImprimer_OnClick;

  if ctxCompta in V_PGI.PGIContexte then
    SetControlVisible('US_CONTROLEUR', True{EstSerie(S7) or EstComptaPackAvance});

{$IFDEF EWS}
  m_bEwsActif := GetParamsocDPSecur ('SO_NE_EWSACTIF', FALSE);
{$ENDIF}

  // $$$ JP 23/05/07: voir les boutons "appeler" si CTI actif
{$IFDEF BUREAU}
  if VH_DP.ctiAlerte <> nil then
  begin
       with GetControl ('BTEL1') as TToolBarButton97 do
       begin
            Visible := TRUE;
            OnClick := AppelerClick;
       end;
       with GetControl ('BTEL2') as TToolBarButton97 do
       begin
            Visible := TRUE;
            OnClick := AppelerClick;
       end;
       with GetControl ('BTEL3') as TToolBarButton97 do
       begin
            Visible := TRUE;
            OnClick := AppelerClick;
       end;
  end;
{$ENDIF}

  // GHA 21/08/07 - Mise à jour de la ComboBox US_GROUPE. suite FQ 11599
  AvertirTable('TTUSERGROUPE');
  
  //PT1 - Début
  // Dans le cas où on gère les intervenants extérieurs, la table intérimaires contient à la fois
  // ces derniers mais aussi les salariés. On peut donc utiliser la tablette correspondante.
  // Sert également dans le cadre du multi-dossier.
  If GetParamsocSecur('SO_PGINTERVENANTEXT', False) Then
  Begin
  	Edit := {$IFNDEF EAGLCLIENT}THDBEdit{$ELSE}THEdit{$ENDIF}(GetControl('US_AUXILIAIRE'));
  	If Edit <> Nil Then Edit.DataType := 'PGSALARIEINT';
  End;
  //PT1 - Fin
end ;

{$IFDEF BUREAU}
procedure TOM_UTILISAT.AppelerClick (Sender:TObject);
var
   strTel :string;
begin
     if VH_DP.ctiAlerte <> nil then
     begin
          strTel := '';
          with Sender as TToolBarButton97 do
          begin
               if Name = 'BTEL1' then
                    strTel := Trim (GetControlText ('US_TEL1'))
               else if Name = 'BTEL2' then
                    strTel := Trim (GetControlText ('US_TEL2'))
               else if Name = 'BTEL3' then
                    strTel := Trim (GetControlText ('US_TEL3'));
          end;

          if (strTel <> '') and (PgiAsk ('Appeler le ' + strTel + ' ?') = mrYes) then
             VH_DP.ctiAlerte.MakeCall (strTel);
     end;
end;

{$ENDIF}

procedure TOM_UTILISAT.ChargeMagUser ;
// recharge les V_PGI dépendants du user
BEGIN
if V_PGI.User = GetField ( 'US_UTILISATEUR' ) then
    BEGIN
    ChargeDefaultEmail;
    V_PGI.User := GetField ( 'US_UTILISATEUR' ) ;
    V_PGI.USerName:= GetField ( 'US_LIBELLE' ) ;
    V_PGI.PassWord:= GetField ( 'US_PASSWORD' ) ;
    V_PGI.Groupe:= GetField ( 'US_GROUPE' ) ;
    V_PGI.Superviseur:= ( GetField ( 'US_SUPERVISEUR' ) = 'X' ) ;
    V_PGI.Controleur:= ( GetField ( 'US_CONTROLEUR' ) = 'X');
    V_PGI.LogUser:= ( GetField ( 'US_SUIVILOG' ) = 'X') ;
    V_PGI.QRCouleur:= ( GetField ( 'US_QRCOULEUR' ) = 'X');
    END ;
END ;

procedure TOM_UTILISAT.Pages_OnChange(Sender: TObject);
var
   visib    :Boolean;
begin
     // $$$ JP 14/09/04 - il faut tenir compte si modif' autorisée
     visib := ((m_bUserModifiable = TRUE) and (TPageControl(GetControl('PAGES')).ActivePageIndex = 0));

     SetControlVisible ('BInsert', visib);
     SetControlVisible ('BDelete', visib);
     if (m_bSeriaAgenda = TRUE) and (m_bUserModifiable = TRUE) and (ds.State <> dsInsert) then
         if TPageControl(GetControl('PAGES')).ActivePageIndex = 3 then
            if TRadioButton (GetControl ('RBUTILVERSAUTRES')).Checked then
               SetDroitsAgendaMode (0)
            else
               SetDroitsAgendaMode (1);
end;

procedure TOM_UTILISAT.BImprimer_OnClick(Sender: TObject);
begin
  SetControlVisible('PW1', False);
  SetControlVisible('PW2', False);
  SetControlVisible('EPW1', False);
  SetControlVisible('EPW2', False);
  SetControlVisible('SPW1', False);
  SetControlVisible('SPW2', False);
  TFFiche(Ecran).BImprimerClick(Sender);
  SetControlVisible('PW1', True);
  SetControlVisible('PW2', True);
  SetControlVisible('EPW1', True);
  SetControlVisible('EPW2', True);
  SetControlVisible('SPW1', True);
  SetControlVisible('SPW2', True);
end;

Initialization
  registerclasses ( [ TOM_UTILISAT ] ) ;
end.



