{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 03/11/2005
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : GUIDE (GUIDE)
Mots clefs ... : TOM;GUIDE
*****************************************************************}
Unit CPGUIDE_TOM ;

Interface

uses StdCtrls,
     Controls,
     Classes, 
{$IFNDEF EAGLCLIENT}
     db,
{$IFNDEF DBXPRESS}dbtables, {$ELSE}uDbxDataSet, {$ENDIF}
     Fiche, 
     FichList,
     Fe_main,
     MenuOLG,
{$else}
     eFiche,
     eFichList,
     MenuOLX,
     MaineAGL, // AGLLanceFiche
{$ENDIF}
{$IFDEF VER150}
     Variants,
{$ENDIF}
     forms, 
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOM,
     Ent1,
     HTB97,
     Dialogs,
     ULibEcriture ,
     Windows,
     UTob ;


Type
  TOM_GUIDE = Class (TOM)
  private
    FQuelGuide              : string ;
    FTypeGuide              : string ;
    FListe                  : THGrid ;    // Grille de saisie
    FMsgBox                 : THMsgBox ;
    BInsert                 : TToolbarButton97;
    BDelete                 : TToolbarButton97;
    FJournal                : THValComboBox;
    FDevise                 : THValComboBox;
    FNaturePiece            : THValComboBox;
    FInfo                   : TInfoEcriture ;
    AKeyDown                : TKeyEvent ;
    FnbLignes               : integer ;
    procedure ChercheGen ;
    procedure ChercheAux ;
    function  GetControlTOF           : boolean; // recupere les contrôles de la grille
    function  AssignEvent             : boolean; // Affecte les evenements au contrôle
    function  InitControl             : boolean; // initialise les grilles
    function  CreateControl           : boolean; // creation de TOB
    procedure InitMsgBox ;
    procedure FormKeyDown ( Sender : TObject ; var Key : Word ; Shift : TShiftState ) ;
    procedure FListeCellEnter(Sender: TObject; var ACol, ARow: Longint;var Cancel: Boolean) ;
    procedure FListeElipsisClick( Sender : TObject ) ;
    procedure FListeKeyPress( Sender : TObject ; var Key : Char);
//    procedure FListeSetEditText( Sender : TObject ; ACol, ARow : Longint; const Value : string) ;
    procedure FListeCellExit( Sender : TObject ; var ACol, ARow : Longint ; var Cancel : Boolean);
    procedure FListeRowEnter(Sender : TObject ; Ou : Integer ; var Cancel : Boolean; Chg : Boolean);
    procedure FListeRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure BVentilClick(Sender: TObject) ;
    procedure BEcheClick( Sender : TObject ) ;
    procedure BInsLigneClick(Sender: TObject);
    procedure BDelLigneClick(Sender: TObject);
    procedure BDupliquerClick (Sender : TObject) ;
    procedure FDeviseChange( Sender : TObject ) ;
    procedure BAssistantClick( Sender : TObject );
    procedure EnableButtons(Ou : integer = -1 ) ;
    procedure CalculeCodeGuide ;
    procedure ValideLeGuide ;
    procedure LoadAna ( vNumLigne : integer ) ;
    procedure VerifGen( Cpte : String ; R : Integer) ;
    function  CtrlGuide : boolean ;
    function  CtrlGuideLigne ( ARow : integer ) : boolean;
    function  LigneAvecArret( ARow : integer ) : boolean ;
    function  CtrlMontant ( Lig : integer ) : boolean;
    procedure SetFoc(C, R: Integer);
    function  CtrlNumeric(St: string; Gene: boolean): boolean;
    function  CtrlAuxiGene ( ARow : integer ) : boolean;
    function  RecupRegle(Gene, Auxi: String): String;
    procedure CreateRow( ARow : integer = -1 ; vBoInsert : boolean = false ) ;
    procedure NextRow ;
    procedure FListeDeleteRow(Row: integer);
    procedure AssisteGuide;
    procedure FListeDblClick(Sender: TObject);
    procedure BGuideClick(Sender: TObject);
    function  CtrlEnTete : boolean;
    function  CtrlCompta : boolean ;
    procedure FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    function  NomGuideOk: boolean;
    procedure FJournalChange( Sender : TObject);
    procedure FormateMontant ( ACol,ARow : Longint ) ;
//    Function  NextCar(ValIn : Char) : Char;
//    Function  NextCode(ValIn : String; taille : integer) : String;
    {JP 16/10/07 : FQ 16149 : Gestion des restrictions établissement}
    procedure GereEtablissement;
  public
    procedure OnNewRecord                 ; override ;
    procedure OnDeleteRecord              ; override ;
    procedure OnUpdateRecord              ; override ;
    procedure OnAfterUpdateRecord         ; override ;
    procedure OnAfterDeleteRecord         ; override ;
    procedure OnLoadRecord                ; override ;
    procedure OnChangeField ( F : TField) ; override ;
    procedure OnArgument ( S : String )   ; override ;
    procedure OnClose                     ; override ;
    procedure OnCancelRecord              ; override ;
  end ;


function ParamGuide ( LeQuel : string ; TypeGuide : String3 ; Mode : TActionFiche ): string ;


Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcGen,
  {$ENDIF MODENT1}
     Lookup,
     GuidUtil,
     GuideAna,
     FichComm,
     ParamSoc,
     SaisUtil,
     {$IFDEF COMPTA}
     Saisie,
     SaisBor,
     {$ENDIF}
     GuidTool;


const GCXGen=1 ; GCGen=2 ;
      GCXAux=3 ; GCAux=4 ;
      GCXRef=5 ; GCRef=6 ;
      GCXLib=7 ; GCLib=8 ;
      GCXDeb=9 ; GCDeb=10;
      GCXCre=11; GCCre=12;
      GCXMrg=13; GCMrg=14;
      GCXEnc=15; GCTva=16;

function ParamGuide ( LeQuel : string ; TypeGuide : String3 ; Mode : TActionFiche ): string ;
begin
 Case Mode of
  taCreat   : result := AGLLanceFiche('CP', 'CPGUIDE', '', TypeGuide+';'+Lequel, 'ACTION=CREATION') ;
  taModif   : result := AGLLanceFiche('CP', 'CPGUIDE', '', TypeGuide+';'+Lequel, 'ACTION=MODIFICATION') ;
  taConsult : result := AGLLanceFiche('CP', 'CPGUIDE', '', TypeGuide+';'+Lequel, 'ACTION=CONSULTATION') ;
 end ;
END ;


procedure TOM_GUIDE.OnNewRecord ;
begin
 inherited ;
 {JP 16/10/07 : FQ 16149 : Gestion des restrictions établissement}
 GereEtablissement;

 Setfield('GU_TYPE',           FTypeGuide) ;
 Setfield('GU_UTILISATEUR',    V_PGI.User) ;
 Setfield('GU_SOCIETE',        V_PGI.CodeSociete) ;
 Setfield('GU_DEVISE',         V_PGI.DevisePivot ) ;
 CalculeCodeGuide ;
 SetControlEnabled('GU_GUIDE',False) ;
 SetFocusControl('GU_LIBELLE') ;
end ;

procedure TOM_GUIDE.OnDeleteRecord ;
var
 i : integer ;
begin

 inherited ;

 for i := 1  to FListe.RowCount - 1 do
  begin
   if FListe.Objects[0,i] <> nil then
    TVentGuide(FListe.Objects[0,i]).Free ;
   FListe.Objects[0,i] := nil ;
  end ;

 ExecuteSQL('DELETE FROM ECRGUI WHERE EG_TYPE="' + FTypeGuide + '" AND EG_GUIDE="' + GetField('GU_GUIDE') + '"') ;
 ExecuteSQL('DELETE FROM ANAGUI WHERE AG_TYPE="' + FTypeGuide + '" AND AG_GUIDE="' + GetField('GU_GUIDE') + '"') ;
 FListe.VidePile(true) ;

 FListe.RowCount := 2 ;

 SetFocusControl('GU_LIBELLE') ;

end ;

procedure TOM_GUIDE.OnUpdateRecord ;
begin
 inherited ;
 Setfield('GU_DATEMODIF',Date) ;
 if not CtrlGuide then
  LastError := 1 ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 05/09/2007
Modifié le ... :   /  /    
Description .. : - LG - 05/09/2007 - Fb 21373 - on affrecte le code retourqd
Suite ........ : lors de la validation
Mots clefs ... : 
*****************************************************************}
procedure TOM_GUIDE.OnAfterUpdateRecord ;
begin
 inherited ;
 ValideLeGuide ;
 TFFiche(Ecran).retour := FQuelGuide ;
 EnableButtons ;
end ;

procedure TOM_GUIDE.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_GUIDE.LoadAna ( vNumLigne : integer ) ;
var
 T   : TVentGuide ;
 QA  : TQuery ;
 St1 : string ;
 St2 : string ;
 i   : integer ;
begin

 if FInfo.LoadCompte(FListe.Cells[GCGen,vNumLigne]) then
 begin
  if FInfo.GetString('G_VENTILABLE') <> 'X' then exit ;
 end
  else
   exit ;

 T := TVentGuide(FListe.Objects[0,vNumLigne]) ;

 if T <> nil then T.Free ;

 T                                  := TVentGuide.Create ;
 FListe.Objects[0,vNumLigne]        := T ;

 QA := OpenSQL('SELECT * FROM ANAGUI WHERE AG_TYPE="' + FTypeGuide + '" AND AG_GUIDE="' + FQuelGuide + '" AND AG_NUMLIGNE='+IntToStr(vNumLigne)+' ORDER BY AG_GUIDE, AG_NUMLIGNE, AG_NUMVENTIL',TRUE) ;

 while not QA.EOF do
  begin
   St1:='' ; St2:=QA.FindField('AG_ARRET').AsString+'              ' ;
   St1:=St1+Format_String(QA.FindField('AG_SECTION').AsString,35)+St2[1] ; Delete(St2,1,1) ;
   St1:=St1+Format_String(QA.FindField('AG_POURCENTAGE').AsString,100)+St2[1] ; Delete(St2,1,1) ;
   St1:=St1+Format_String(QA.FindField('AG_POURCENTQTE1').AsString,100)+St2[1] ; Delete(St2,1,1) ;
   St1:=St1+Format_String(QA.FindField('AG_POURCENTQTE2').AsString,100)+St2[1] ; Delete(St2,1,1) ;
   i:=StrToInt(Copy(QA.FindField('AG_AXE').AsString,2,1)) ;
   T.Ventil[i].Add(St1) ;
   QA.Next;
  end ;

 Ferme(QA) ;

end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 28/04/2006
Modifié le ... :   /  /    
Description .. : - LG - FB 17959 - 28/04/2006 - mise a jour de la nature de
Suite ........ : piece
Mots clefs ... : 
*****************************************************************}
procedure TOM_GUIDE.OnLoadRecord ;
var
 Q  : TQuery ;
 St : string ;
begin

 Inherited ;

 FQuelGuide := GetField('GU_GUIDE') ;

 FListe.VidePile(true) ;
 FListe.RowCount := 2 ;
 FnbLignes := 0 ;

 Q := OpenSQL('SELECT * FROM ECRGUI WHERE EG_TYPE="' + FTypeGuide + '" AND EG_GUIDE="' + FQuelGuide + '" ORDER BY EG_GUIDE, EG_NUMLIGNE',TRUE) ;
 while not Q.EOF do
  begin

   St        := Q.FindField('EG_ARRET').AsString+'                       ' ;
   FnbLignes := Q.FindField('EG_NUMLIGNE').AsInteger ;

   FListe.Cells[0,FListe.RowCount-1]             := IntToStr(FnbLignes);

   FListe.Cells[GCGen,FListe.RowCount-1]         := Q.FindField('EG_GENERAL').AsString ;
   FListe.Cells[GCXGen,FListe.RowCount-1]        := St[1] ; Delete(St,1,1) ;

   FListe.Cells[GCAux,FListe.RowCount-1]         :=Q.FindField('EG_AUXILIAIRE').AsString ;
   FListe.Cells[GCXAux,FListe.RowCount-1]        :=St[1] ; Delete(St,1,1) ;

   FListe.Cells[GCRef,FListe.RowCount-1]         :=Q.FindField('EG_REFINTERNE').AsString ;
   FListe.Cells[GCXRef,FListe.RowCount-1]        :=St[1] ; Delete(St,1,1) ;

   FListe.Cells[GCLib,FListe.RowCount-1]         :=Q.FindField('EG_LIBELLE').AsString ;
   FListe.Cells[GCXLib,FListe.RowCount-1]        :=St[1] ; Delete(St,1,1) ;

   FListe.Cells[GCDeb,FListe.RowCount-1]         :=Q.FindField('EG_DEBITDEV').AsString ;
   FListe.Cells[GCXDeb,FListe.RowCount-1]        :=St[1] ; Delete(St,1,1) ;

   FListe.Cells[GCCre,FListe.RowCount-1]         :=Q.FindField('EG_CREDITDEV').AsString ;
   FListe.Cells[GCXCre,FListe.RowCount-1]        :=St[1] ; Delete(St,1,1) ;

   FListe.CellValues[GCMrg,FListe.RowCount-1]    :=Q.FindField('EG_MODEREGLE').AsString ;
   FListe.Cells[GCXMrg,FListe.RowCount-1]        :=St[1] ; Delete(St,1,1) ;

   if VH^.OuiTvaEnc then
    begin
     FListe.Cells[GCXEnc,FListe.RowCount-1]      :=Q.FindField('EG_TVAENCAIS').AsString ;
     FListe.CellValues[GCTva,FListe.RowCount-1]  :=Q.FindField('EG_TVA').AsString ;
    end ;

   LoadAna(FnbLignes) ;
   FListe.RowCount                               := FListe.RowCount + 1 ;
   Q.Next ;

  end ; // while

 Ferme(Q) ;
 if FListe.RowCount > 2 then FListe.RowCount := FListe.RowCount - 1 ;

 Ecran.Caption := TraduireMemoire('Guides de saisie') + ' : ' + GetControlText('GU_LIBELLE') ;
 UpdateCaption(Ecran) ;

 FDeviseChange(nil) ;
 FJournalChange(nil) ;

end ;

procedure TOM_GUIDE.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 18/10/2007
Modifié le ... :   /  /    
Description .. : - LG - 18/10/2007 - FB 21136 - on posotionne le focus sur 
Suite ........ : le libelle a l'entree de la fiche
Mots clefs ... : 
*****************************************************************}
procedure TOM_GUIDE.OnArgument ( S : String ) ;
var
 lSt : string ;
begin

 Inherited ;

 lSt         := TFFiche(Ecran).FLequel ;
 FTypeGuide  := readtokenSt(lSt) ;
 FQuelGuide  := readtokenSt(lSt) ;

 if not GetControlTOF then exit ;
 if not CreateControl then exit ;
 if not InitControl   then exit ;
 if not AssignEvent   then exit ;

 SetFocusControl('GU_LIBELLE') ;

end ;

procedure TOM_GUIDE.OnClose ;
begin

 FListe.VidePile(true) ;
 FMsgBox.Free ;
 FInfo.Free ;

 inherited ;
 
end ;

procedure TOM_GUIDE.OnCancelRecord ;
begin
  Inherited ;
end ;

function TOM_GUIDE.AssignEvent : boolean;
begin
 AKeyDown                                                  := Ecran.OnKeyDown ;
 Ecran.OnKeyDown                                           := FormKeyDown ;
 FListe.OnCellEnter                                        := FListeCellEnter ;
 FListe.OnElipsisClick                                     := FListeElipsisClick ;
 FListe.OnKeyPress                                         := FListeKeyPress ;
// FListe.OnSetEditText                                      := FListeSetEditText ;
 FListe.OnCellExit                                         := FListeCellExit ;
 FListe.OnRowEnter                                         := FListeRowEnter ;
 FListe.OnRowExit                                          := FListeRowExit ;
 FListe.OnDblClick                                         := FListeDblClick ;
 FListe.OnKeyDown                                          := FListeKeyDown ;
 FDevise.OnChange                                          := FDeviseChange ;
 FJournal.OnChange                                         := FJournalChange ;
 TToolbarButton97(GetControl('BVentil',true)).OnClick      := BVentilClick ;
 TToolbarButton97(GetControl('BEche',true)).OnClick        := BEcheClick ;
 TToolbarButton97(GetControl('BInsLigne',true)).OnClick    := BInsLigneClick ;
 TToolbarButton97(GetControl('BDelLigne',true)).OnClick    := BDelLigneClick ;
 TToolbarButton97(GetControl('BGuide',true)).OnClick       := BGuideClick ;
 TToolbarButton97(GetControl('BDupliquer',true)).OnClick   := BDupliquerClick ;
 TToolbarButton97(GetControl('BAssistant',true)).OnClick  := BAssistantClick ;
 result                                                    := true ;
end;

function TOM_GUIDE.CreateControl : boolean;
begin
 FMsgBox    := THMsgBox.create(FMenuG) ;
 FInfo      := TInfoEcriture.Create() ;
 result     := true ;
end;

function TOM_GUIDE.GetControlTOF : boolean;
begin
 FListe         := THGrid(GetControl('FLISTE',true)) ;
 BInsert        := TToolbarButton97(GetControl('BInsert',true)) ;
 BDelete        := TToolbarButton97(GetControl('BDelete',true)) ;
 FJournal       := THValComboBox(GetControl('GU_JOURNAL',true));
 FDevise        := THValComboBox(GetControl('GU_DEVISE',true));
 FNaturePiece   := THValComboBox(GetControl('GU_NATUREPIECE',true));

 result         := true ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : LG
Créé le ...... : 24/05/2007
Modifié le ... :   /  /    
Description .. : - LG - FB 19192 - voir fiche
Mots clefs ... : 
*****************************************************************}
function TOM_GUIDE.InitControl : boolean;
begin

 InitMsgBox ;

 FListe.ColFormats[GCMrg] := 'CB=TTMODEREGLE||' ;
 FListe.ColFormats[GCTva] := 'CB=TTTVA||' ;

 if (( not VH^.OuiTvaEnc) or ( FTypeGuide = 'ENC') or ( FTypeGuide = 'DEC')) then
  begin
   FListe.ColCount := FListe.ColCount - 2 ;
   //Fliste.ColWidths[GCXEnc] := -1 ;
   //Fliste.ColWidths[GCTva]  := -1 ;
  end ;

 SetControlVisible('GU_TRESORERIE', FTypeGuide = 'NOR' ) ;
 SetControlProperty('GU_TRESORERIE','AllowGrayed',False) ; 
 FJournal.DataType := 'TTJALSAISIE' ;
 // GP le 20/08/2008 21496
 RetoucheHVCPoursaisie(FJournal) ;


 if FTypeGuide = 'ABO' then
  Ecran.Caption := FMsgBox.Mess[22]
   else
    if ( (FTypeGuide ='ENC') or (FTypeGuide='DEC') ) then
     begin
      Ecran.Caption     := FMsgBox.Mess[23] ;
      BInsert.Enabled   := False ;
      BDelete.Enabled   := False ;
     end
      else
       begin
        FJournal.DataType := 'ttJalGuide' ;
        Ecran.Caption     := FMsgBox.Mess[21] ;
       end ;

 TToolbarButton97(GetControl('BVentil',true)).Visible := not VH^.AnaCroisaxe ;

 {$IFNDEF COMPTA}
  TToolbarButton97(GetControl('BGUIDE',true)).Visible := False;
 {$ENDIF}

 if ( ctxPCL in V_PGI.PGIContexte ) then
  TToolbarButton97(GetControl('BVentil',true)).Visible := ( GetParamSocSecur('SO_CPPCLSANSANA',False)= true ) and not VH^.AnaCroisaxe
   else
    TToolbarButton97(GetControl('BVentil',true)).Visible := not VH^.AnaCroisaxe;

 {JP 19/06/06 : c'est beaucoup mieux ici que dans FJournalChange}
 SetControlEnabled('GU_DEVISE', FTypeGuide <> 'ABO');
 
 result := true ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 25/04/2006
Modifié le ... :   /  /
Description .. : - LG - 25/04/2006 - FB 17363 - changement du message 
Suite ........ : 13
Mots clefs ... : 
*****************************************************************}
procedure TOM_GUIDE.InitMsgBox ;
begin
 FMsgBox.Mess.Add('0;Guides de saisie; Compte général fermé : vous ne pouvez pas saisir d''écritures sur ce compte.;W;O;O;O;') ;
 FMsgBox.Mess.Add('1;Guides de saisie; Compte auxiliaire fermé : vous ne pouvez pas saisir d''écritures sur ce compte.;W;O;O;O;') ;
 FMsgBox.Mess.Add('2;Guides de saisie; Compte général inexistant : vous devez impérativement lui associer un arrêt en saisie.;W;O;O;O;' ) ;
 FMsgBox.Mess.Add('3;Guides de saisie; Compte auxiliaire inexistant : vous devez impérativement lui associer un arrêt en saisie.;W;O;O;O;' ) ;
 FMsgBox.Mess.Add('4;Guides de saisie; Comptes général et auxiliaire non renseignés : vous devez paramétrer au moins un arrêt en saisie sur ces comptes.;W;O;O;O;' ) ;
 FMsgBox.Mess.Add('5;Guides de saisie; Ce compte général n''est pas collectif : vous ne pouvez ni lui associer un compte auxiliaire, ni paramétrer un arrêt en saisie sur le compte auxiliaire.;W;O;O;O;' ) ;
 FMsgBox.Mess.Add('6;Guides de saisie; Ce compte général est collectif : vous devez lui associer un compte auxiliaire, ou paramétrer un arrêt en saisie sur le compte auxiliaire.;W;O;O;O;' ) ;
 FMsgBox.Mess.Add('7;Guides de saisie; Vous ne pouvez pas renseigner simultanément les zones Débit et Crédit.;W;O;O;O;' ) ;
 FMsgBox.Mess.Add('8;Guides de saisie; Le renseignement d''un montant interdit l''arrêt en saisie sur les zones de montant.;W;O;O;O;' ) ;
 FMsgBox.Mess.Add('9;Guides de saisie; La valeur saisie dans la zone de compte général doit être une chaîne de caractères numériques ou une formule.;W;O;O;O;' ) ;
 FMsgBox.Mess.Add('10;Guides de saisie; La valeur saisie dans les zones de montant doit être une chaîne de caractères numériques ou une formule.;W;O;O;O;' ) ;
 FMsgBox.Mess.Add('11;Guides de saisie; Les montants négatifs ne sont pas autorisés sur cette société.;W;O;O;O;' ) ;
 FMsgBox.Mess.Add('12;Guides de saisie; Erreur de syntaxe dans la formule.;W;O;O;O;' ) ;
 FMsgBox.Mess.Add('13;Guides de saisie; Erreur dans la formule: mauvaise utilisation du caractère '':'' .;W;O;O;O;' ) ;
 FMsgBox.Mess.Add('14;Guides de saisie; Ce journal n''est pas multi-devise, il est incompatible avec la devise choisie.;W;O;O;O;' ) ;
 FMsgBox.Mess.Add('15;Guides de saisie; Erreur dans la formule: référence à un champ inexistant.;W;O;O;O;' ) ;
 FMsgBox.Mess.Add('16;Guides de saisie; Ce compte n''est pas ventilable !;W;O;O;O;' ) ;
 FMsgBox.Mess.Add('17;Guides de saisie; Vous devez renseigner un montant ou bien paramétrer un arrêt en saisie sur les montants.;W;O;O;O;' ) ;
 FMsgBox.Mess.Add('18;Suppression; Confirmez-vous la suppression de ce guide ?;Q;YNC;N;N;' ) ;
 FMsgBox.Mess.Add('19;Guides de saisie; Voulez-vous enregistrer les modifications ?;Q;YNC;N;N;' ) ;
 FMsgBox.Mess.Add('20;Guides de saisie; Vous devez indiquer le nom du guide avant de l''enregistrer ;W;O;O;O;' ) ;
 FMsgBox.Mess.Add('Guides de saisie' ) ;
 FMsgBox.Mess.Add('Guides d''abonnement' ) ;
 FMsgBox.Mess.Add('Guides de règlement' ) ;
 FMsgBox.Mess.Add('Guides de pointage' ) ;
 FMsgBox.Mess.Add('25;Guides de saisie;Cette ligne est vide. Vous devez renseigner au minimum la référence ou le libellé.;W;O;O;O;' ) ;
 FMsgBox.Mess.Add('26;Guides de saisie;Vous devez indiquer un montant ;W;O;O;O;' ) ;
 FMsgBox.Mess.Add('27;Guides de saisie;Vous devez indiquer un compte existant.;W;O;O;O;' ) ;
 FMsgBox.Mess.Add('28;Guides de saisie;Ce guide est vide ! ;W;O;O;O;' ) ;
 FMsgBox.Mess.Add('29;Guides de saisie;Vous devez indiquer le journal de saisie;W;O;O;O;' ) ;
 FMsgBox.Mess.Add('30;Guides de saisie;Vous devez indiquer une nature de pièce;W;O;O;O;' ) ;
 FMsgBox.Mess.Add('31;Guides de saisie;Vous devez indiquer une devise;W;O;O;O;' ) ;
 FMsgBox.Mess.Add('32;Guides de saisie;Guide utilisé pour la génération d''abonnements : la pièce doit être équilibrée;W;O;O;O;' ) ;
 FMsgBox.Mess.Add('33;Guides de saisie;Attention : cette ligne est ventilée sur des sections analytiques, souhaitez-vous maintenir la ventilation?;Q;YN;N;N;') ;
 FMsgBox.Mess.Add('Enregistrement non validé !' ) ;
 FMsgBox.Mess.Add('35;Guides de saisie;Vous devez renseigner une devise !;E;O;O;O;' ) ;
 FMsgBox.Mess.Add('36;Guides de saisie;Vous devez renseigner un journal avant de tester le guide !;E;O;O;O;' ) ;
 FMsgBox.Mess.Add('37;Guides de saisie;Vous devez renseigner un journal avec des compteurs valides !;E;O;O;O;' ) ;
 FMsgBox.Mess.Add('38;Guides de saisie;Vous ne pouvez pas renseigner des montants négatifs !;E;O;O;O;' ) ;
 FMsgBox.Mess.Add('39;Guides de saisie;Vous devez renseigner un journal !;E;O;O;O;' ) ;
 FMsgBox.Mess.Add('40;Guides de saisie;Vous devez renseigner un établissement !;W;O;O;O;' ) ;
 FMsgBox.Mess.Add('41;Guides de saisie;Ce nom de guide existe déjà !;W;O;O;O;' ) ;
 FMsgBox.Mess.Add('42;Guides de saisie; La nature du compte général est incompatible avec la nature de la pièce.;W;O;O;O;' ) ;
 FMsgBox.Mess.Add('43;Guides de saisie; La nature du compte auxiliaire est incompatible avec la nature de son collectif.;W;O;O;O;' ) ;
 FMsgBox.Mess.Add('44;Guides de saisie; Vous devez saisir au moins une ligne contenant le compte de contrepartie du journal saisi !;W;O;O;O;' ) ;
 FMsgBox.Mess.Add('EURO') ;
 FMsgBox.Mess.Add('46;Guides de saisie; Vous ne pouvez pas choisir un journal dédié aux tiers payeurs;W;O;O;O;' ) ;
 FMsgBox.Mess.Add('47;Guides de saisie; Vous ne pouvez pas saisir une formule sur plus de 100 caractères !;W;O;O;O;' ) ;
 FMsgBox.Mess.Add('48;Guides de saisie; Vous devez enregistrer le guide en cours;W;O;O;O;' ) ;

end;

procedure TOM_GUIDE.ChercheGen ;
var
 lStCompte     : string ;
 lStSelect     : string ;
 lStColonne    : string ;
 lStOrder      : string ;
 lStWhere      : string ;
 lBoResult     : boolean ;
begin

 CMakeSQLLookupGen(lStWhere,lStColonne,lStOrder,lStSelect) ;
 lBoResult := LookupList(FListe,TraduireMemoire('Comptes'),'GENERAUX',lStColonne,lStSelect,lStWhere,lStOrder,true, 1,'',tlLocate) ;

 if not lBoResult then exit ;
 
 lStCompte := uppercase(FListe.Cells[GCGen,FListe.Row]) ;
 FInfo.Journal.Load([GetField('GU_JOURNAL')] ) ;

 if FInfo.LoadCompte(lStCompte) then
  begin

   if EstInterdit( FInfo.Journal.GetValue('J_COMPTEINTERDIT') , FInfo.StCompte , 0 ) > 0 then
    begin
     FListe.Cells[GCGen,FListe.Row] := '';
     PgiBox('Ce compte général est interdit sur ce journal.', Ecran.Caption);
     exit;
    end ;

   if ( FInfo.GetString('G_FERME') = 'X' ) then
    begin
     FListe.Cells[GCGen,FListe.Row] := '';
     PgiBox('Ce compte est fermé. Vous ne pouvez plus l''utiliser dans un guide.', Ecran.Caption);
     exit;
    end ;

  end ; // if

 // compte OK
 FListe.Cells[GCGen,FListe.Row] := lStCompte ;

end;

procedure TOM_GUIDE.ChercheAux ;
var
 lStSelect     : string ;
 lStColonne    : string ;
 lStOrder      : string ;
 lStWhere      : string ;
 lStNatureGene : string ;
begin

 lStNatureGene :=  FInfo.GetString('G_NATUREGENE') ;

 CMakeSQLLookupAux(lStWhere,lStColonne,lStOrder,lStSelect,GetField('GU_NATRUEPIECE'),lStNatureGene) ;
 LookupList(FListe,TraduireMemoire('Auxiliaire'),'TIERS',lStColonne,lStSelect,lStWhere,lStOrder,true, 2,'',tlLocate) ;

end;

procedure TOM_GUIDE.FormKeyDown(Sender: TObject; var Key: Word;Shift: TShiftState);
var
 Vide : boolean ;
begin

 Vide := (Shift = []) ;

 AKeyDown(Sender,Key,Shift) ;

 case Key of
  VK_F5  : if Vide then
               if (Screen.ActiveControl=FListe) then
                begin
                 Key:=0 ;
                  case FListe.Col of
                   GCGen : ChercheGen ;
                   GCAux : ChercheAux ;
                   GCRef,
                   GCLib,
                   GCDeb,
                   GCCre : BAssistantClick(nil) ;
                  end ; // case
                end
                 else
                  begin
                   Key:=0 ;
                   AssisteGuide ;
                  end ;
  VK_ESCAPE : if Vide then
                begin
                 Key := 0 ;
                 Ecran.Close ;
                end ;
  VK_TAB     : if Vide then
               begin
                if ( FListe.Row = FListe.RowCount - 1 ) and ( FListe.Col = FListe.Colcount - 1 ) then
                  begin
                   key := 0 ;
                   NextRow ;
                  end;
               end;  // if
  VK_DOWN   : if ( FListe.InplaceEditor.Focused  or FListe.Focused ) and ( FListe.Row = FListe.RowCount - 1 )  then
               begin
                Key := 0 ;
                NextRow ;
               end;
  VK_INSERT : if (Vide) then
               begin
                Key := 0 ;
                CreateRow(FListe.Row,true) ;
               end; // if
  VK_RETURN  : if (Vide) then
               begin
                if ( FListe.Col = FListe.ColCount -1 ) and ( FListe.Row = FListe.RowCount -1 ) then
                begin
                 Key := 0 ; NextRow ;
                end
                 else
                  Key := VK_TAB ;
             end ;
 VK_DELETE : if  Shift=[ssCtrl] then
              begin
               Key := 0 ;
               BDelLigneClick(nil) ;
              end ;
{AA}    65 : if ((Shift=[ssAlt]) and (TToolbarButton97(GetControl('BVentil',true)).Enabled)) then begin Key:=0 ; BVentilClick(nil) ; end ;
{AE}    69 : if Shift=[ssAlt] then begin Key:=0 ; BEcheClick(nil) ; end ;
{AG}    71 : if Shift=[ssAlt] then begin Key:=0 ; BGuideClick(nil) ; end ;
{^Z}    90 :  if Shift=[ssAlt] then begin Key:=0 ; if Ecran.WindowState=wsMaximized then Ecran.WindowState:=wsNormal else Ecran.WindowState:=wsMaximized end ;

    end ;
end;

procedure TOM_GUIDE.FListeCellEnter(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean) ;
var
 lc : integer ;
begin

 if ((Not VH^.OuiTvaEnc) or ( FTypeGuide = 'ENC') or ( FTypeGuide = 'DEC') ) then
  Cancel := ((FListe.Col=GCXEnc) or (FListe.Col=GCTva)) ;

 lc := FListe.Col ;

 FListe.ElipsisButton := ( lc =  GCGen ) or ( lc =  GCAux ) ;

end;


procedure TOM_GUIDE.FListeKeyPress( Sender : TObject ; var Key: Char);
begin

 if ( Key <> #9 ) and ( Key <> #0 ) then
  begin
    if DS.State = dsBrowse then
     DS.Edit ;
    // on force la maj d'un champs pour passer en modification
    Setfield('GU_GUIDE', GetField('GU_GUIDE')) ;
  end ; // if

end;

procedure TOM_GUIDE.FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
if ( key = VK_DELETE ) then
 begin
  if DS.State = dsBrowse then
   DS.Edit ;
  // on force la maj d'un champs pour passer en modification 
  Setfield('GU_GUIDE', GetField('GU_GUIDE')) ;
 end ;
end;

(*
procedure TOM_GUIDE.FListeSetEditText( Sender : TObject ; ACol, ARow : Longint; const Value : string) ;
begin

 //if DS.State = dsBrowse then
//  DS.Edit ;
{
if ((FListe.Row=FListe.RowCount-1) and (Not G_LigneVide(FListe,FListe.Row))) then
   BEGIN
   if ((TypeGuide='ENC') or (TypeGuide='DEC')) and (FListe.RowCount=3) then else FListe.RowCount:=FListe.RowCount+1 ;
   END;
if ((FListe.Row=FListe.RowCount-3) and (G_LigneVide(FListe,FListe.Rowcount-1))
                                   and (G_LigneVide(FListe,FListe.Rowcount-2))) then
   BEGIN
   if FListe.Objects[0,FListe.Rowcount-1]<>NIL then TVentGuide(FListe.Objects[0,FListe.Rowcount-1]).Free ;
   FListe.RowCount:=FListe.RowCount-1 ;
   END;  }
end; *)

procedure TOM_GUIDE.FormateMontant ( ACol,ARow : Longint ) ;
var
 X : Double ;
 St : String ;
begin
 St := FListe.Cells[ACol,ARow] ;
 if St='' then Exit ;
 X  := Valeur(St) ;
 if X = 0 then Exit ;
 if Not IsNumeric(St) then Exit ;
 St:=StrfMontant(X,15,FInfo.Devise.Dev.Decimale,'',True) ;
 FListe.Cells[ACol,ARow] := St ;
end ;


procedure TOM_GUIDE.FListeCellExit( Sender : TObject ; var ACol, ARow : Longint ; var Cancel : Boolean);
var
 St : string ;
begin

 EnableButtons ;

 St :=  FListe.Cells[ACol,ARow];

 if Length(FListe.Cells[ACol,ARow]) > 100 then
  begin
   FMsgBox.Execute(47,'','') ;
   FListe.Cells[ACol,ARow] := Copy(FListe.Cells[ACol,ARow],1,100) ;
   Cancel:=True ;
   Exit ;
  end ;

  case ACol of
  GCGen   : VerifGen(FListe.Cells[GCGen,ARow],ARow) ;
  GCXGen,
  GCXAux,
  GCXRef,
  GCXLib,
  GCXDeb,
  GCXCre,
  GCXMrg  : if UpperCase(FListe.Cells[Acol,ARow]) <> 'X' then
             FListe.Cells[Acol,ARow]:='-' ;
  GCXEnc  : if ((FListe.Cells[Acol,ARow]<>'X') and (FListe.Cells[Acol,ARow]<>'-')) then
             FListe.Cells[Acol,ARow]:='' ;
  GCDeb,
  GCCre   : begin
             FormateMontant(ACol,ARow) ;
            // if Valeur(FListe.Cells[ACol,ARow]) = 0 then Exit ;
            // FListe.Cells[ACol,ARow] := StrfMontant( Valeur(FListe.Cells[ACol,ARow]) ,15,FInfo.Devise.Dev.Decimale,'',True) ;
            end ;
 end ; // case

end;


procedure TOM_GUIDE.FListeElipsisClick(Sender: TObject);
begin

 if csDestroying in Ecran.ComponentState then Exit ;

 if FListe.Col = GCGen then
  ChercheGen
   else
    if FListe.Col = GCAux then
     ChercheAux
      else
       if FListe.Col = GCMrg then
        LookupList(FListe,TraduireMemoire('Mode de règlement'),'TTMODEREGLE','','','','',true, 0,'',tlLocate) ;

 if DS.State = dsBrowse then
  DS.Edit ;

end;

procedure TOM_GUIDE.FListeRowEnter( Sender : TObject ; Ou : Longint ; var Cancel : Boolean ; Chg : Boolean) ;
begin
 EnableButtons() ;
end;

procedure TOM_GUIDE.FListeRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin

end ;

procedure TOM_GUIDE.EnableButtons( Ou : integer = -1 ) ;
var
 lStCompte : string ;
begin

 if Ou = - 1 then Ou := Fliste.Row ;

 lStCompte := Uppercase( FListe.Cells[GCGen,Ou]) ;

 if FInfo.LoadCompte(lStCompte) then
  begin
   SetControlVisible('BVENTIL', FInfo.GetString('G_VENTILABLE') = 'X' ) ;
  end ;

 TToolbarButton97(GetControl('BDupliquer',true)).Enabled := DS.State = dsBrowse ;

end ;


procedure TOM_GUIDE.CalculeCodeGuide ;
//var
// lQ : TQuery ;
begin

 if DS.State in [dsInsert] then
  begin
  {  lQ := OpenSQL('Select MAX(GU_GUIDE) from GUIDE Where GU_TYPE="' + FTypeGuide + '" Order by 1',True) ;
    if not lQ.EOF then
     begin
      FQuelGuide := lQ.Fields[0].AsString ;
      if FQuelGuide <> '' then
        FQuelGuide:=NextCode(FQuelGuide, 3)  //YMO 09/10/06 Dev 4626 Code alphanumérique (ex : guides d'abo > 999)
      else
        FQuelGuide := '001' ;

      while Length(FQuelGuide) < 3 do
         FQuelGuide := '0' + FQuelGuide ;

     end
      else
       FQuelGuide := '001' ;   }
    FQuelGuide := CCalculeCodeGuide(FTypeguide) ;  

    SetField('GU_GUIDE', FQuelGuide) ;
    SetControlText('GU_GUIDE', FQuelGuide) ;

  end ;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 09/10/2006
Modifié le ... : 16/10/2006
Description .. : Incrémentation d'une valeur sur 36 positions
Mots clefs ... :
*****************************************************************}
{function TOM_GUIDE.NextCar(ValIn : Char) : Char;
begin

  if Ord(ValIn)=57 then
      Result:=Chr(65)   // 9 --> A
  else
    if Ord(ValIn)<>90 then
        Result:=Chr(Ord(ValIn)+1)
    else
        Result:=Chr(48);  // Z --> 0

end; }

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 09/10/2006
Modifié le ... : 16/10/2006
Description .. : Incrémentation d'une chaîne sur 36 positions
Mots clefs ... :
*****************************************************************}
{Function TOM_GUIDE.NextCode(ValIn : String; taille : integer) : String;
var
  ValOut : String ;
  i, j  : integer ;
  mov : boolean;
begin

 ValOut:='';
 ValIn:=Uppercase(ValIn+'Z');

 for i:=1 to taille do
 begin
      mov:=true;
      for j:=i+1 to taille+1 do      // si tous les suivants à Z...
          If ValIn[j]<>'Z' then mov:=false;

      If mov then
        ValOut:=ValOut+NextCar(ValIn[i]) // ...on incrémente
      else
        ValOut:=ValOut+ValIn[i];
 end;

 Result:=ValOut;

end;}


procedure TOM_GUIDE.ValideLeGuide ;
var
 lQ         : TQuery ;
 St         : string ;
 St1        : string ;
 i          : integer ;
 nb         : integer ;
 ix         : integer ;
 ia         : integer ;
 TEcrGuide  : TOB ;
 TAnaGuide  : TOB ;
 T          : TVentGuide ;

 procedure _SetG_Croix( Value : string ) ;
 begin
  if G_Croix(Value) then
   St := St + 'X'
    else
     St := St + '-' ;
 end ;

 procedure _DeleteSt( Value : integer ) ;
 begin
  Delete(St1,1,Value) ; St:=St+St1[1] ; Delete(St1,1,1) ;
 end ;

begin

  // On supprime pour ré-insérer
  ExecuteSQL('DELETE FROM ECRGUI WHERE EG_TYPE="' + FTypeGuide + '" AND EG_GUIDE="' + FQuelGuide + '"') ;
  ExecuteSQL('DELETE FROM ANAGUI WHERE AG_TYPE="' + FTypeGuide + '" AND AG_GUIDE="' + FQuelGuide + '"') ;

  CalculeCodeGuide ;

  nb := 0 ;


  for i:=1 to FListe.RowCount - 1 do
    begin
     Inc(nb) ;
     St        := '' ;
     TEcrGuide := TOB.Create('ECRGUI',nil,-1) ;
     TEcrGuide.InitValeurs ;
     TEcrGuide.PutValue('EG_TYPE'       , FTypeGuide ) ;
     TEcrGuide.PutValue('EG_GUIDE'      , FQuelGuide ) ;
     TEcrGuide.PutValue('EG_NUMLIGNE'   , nb ) ;
     TEcrGuide.PutValue('EG_GENERAL'    , FListe.Cells[GCGen,i] ) ;
     TEcrGuide.PutValue('EG_AUXILIAIRE' , FListe.Cells[GCAux,i]) ;
     TEcrGuide.PutValue('EG_REFINTERNE' , FListe.Cells[GCRef,i] ) ;
     TEcrGuide.PutValue('EG_LIBELLE'    , FListe.Cells[GCLib,i] ) ;
     TEcrGuide.PutValue('EG_DEBITDEV'   , FListe.Cells[GCDeb,i] ) ;
     TEcrGuide.PutValue('EG_CREDITDEV'  , FListe.Cells[GCCre,i] ) ;
     TEcrGuide.PutValue('EG_MODEREGLE'  , FListe.CellValues[GCMrg,i]) ;

     if VH^.OuiTvaEnc then
      begin
       St := FListe.Cells[GCXEnc,i] ;
       if ( ( St <> 'X' ) and ( St <> '-' )) then
        St := '' ;
       TEcrGuide.PutValue('EG_TVAENCAIS', St) ;
       TEcrGuide.PutValue('EG_TVA'      , FListe.CellValues[GCTva,i] ) ;
      end ;

     // FQ 12090
     if ( FListe.Cells[GCAux,i] = '') then
      TEcrGuide.PutValue('EG_RIB', '' )
       else
        begin
         lQ := OpenSQL('SELECT R_CODEIBAN, R_PAYS, R_ETABBQ, R_GUICHET, R_NUMEROCOMPTE, R_CLERIB, R_DOMICILIATION FROM RIB WHERE R_AUXILIAIRE="' + FListe.Cells[GCAux,i] + '" AND R_PRINCIPAL="X"', True) ;
         if lQ.EOF then
          TEcrGuide.PutValue('EG_RIB', '')
           else
            begin
             if (lQ.FindField('R_CODEIBAN').AsString = '') then
              TEcrGuide.PutValue('EG_RIB', calcRIB(lQ.FindField('R_PAYS').AsString, lQ.FindField('R_ETABBQ').AsString, lQ.FindField('R_GUICHET').AsString, lQ.FindField('R_NUMEROCOMPTE').AsString, lQ.FindField('R_CLERIB').AsString))
               else
                TEcrGuide.PutValue('EG_RIB', lQ.FindField('R_CODEIBAN').AsString);
            end;
        end;

     St :=  '' ;
     _SetG_Croix(FListe.Cells[GCXGen,i]) ;
     _SetG_Croix(FListe.Cells[GCXAux,i]) ;
     _SetG_Croix(FListe.Cells[GCXRef,i]) ;
     _SetG_Croix(FListe.Cells[GCXLib,i]) ;
     _SetG_Croix(FListe.Cells[GCXDeb,i]) ;
     _SetG_Croix(FListe.Cells[GCXCre,i]) ;
     _SetG_Croix(FListe.Cells[GCXMrg,i]) ;
     TEcrGuide.PutValue('EG_ARRET', St) ;

     TEcrGuide.InsertDB(nil) ;
     TEcrGuide.Free ;

     T := TVentGuide(FListe.Objects[0,i]) ;
     if T = nil then continue ;

     for ix := 1 to MaxAxe do
      for ia := 0 to T.Ventil[ix].Count - 1 do
       if trim(T.Ventil[ix].Strings[ia])<>'' then
        begin
         St1       := T.Ventil[ix].Strings[ia] ;
         St        := '' ;
         TAnaGuide := TOB.Create('ANAGUI',nil,-1) ;
         TAnaguide.InitValeurs ;
         TAnaGuide.PutValue('AG_TYPE'           , FTypeGuide) ;
         TAnaGuide.PutValue('AG_GUIDE'          , FQuelGuide) ;
         TAnaGuide.PutValue('AG_NUMLIGNE'       , nb) ;
         TAnaGuide.PutValue('AG_NUMVENTIL'      , ia+1) ;
         TAnaGuide.PutValue('AG_AXE'            , 'A'+IntToStr(ix)) ;
         TAnaGuide.PutValue('AG_SECTION'        , Trim(Copy(St1,1,35))) ;
         _DeleteSt(35) ;
         TAnaGuide.PutValue('AG_POURCENTAGE'    , Trim(Copy(St1,1,100)) ) ;
         _DeleteSt(100) ;
         TAnaGuide.PutValue('AG_POURCENTQTE1'   , Trim(Copy(St1,1,100)) ) ;
         _DeleteSt(100) ;
         TAnaGuide.PutValue('AG_POURCENTQTE2'   , Trim(Copy(St1,1,100)) ) ;
         _DeleteSt(100) ;
         TAnaGuide.PutValue('AG_ARRET'          , St1 ) ;
         TAnaGuide.InsertDB(nil) ;
         TAnaGuide.Free ;
        end ;

    end ;

end;

procedure TOM_GUIDE.VerifGen( Cpte : String ; R : Integer) ;
var
 iaxe,isec    : integer;
 lBoExistV    : boolean;
 T            : TVentGuide;
 len          : integer ;
begin

 if DS.State = dsBrowse then exit ;

 len := length(FListe.Cells[GCGen,FListe.Row]) ;

 if ( len > 0 ) and ( len < VH^.Cpta[fbGene].Lg ) then
  FListe.Cells[GCXGen,FListe.Row]:='X'
   else
    if UpperCase(FListe.Cells[GCXGen,FListe.Row]) <> 'X' then
     FListe.Cells[GCXGen,FListe.Row]:='-' ;

 lBoExistV  := false;

 T          := TVentGuide(FListe.Objects[0,R]) ;

 if T = nil then exit ;

// GP le 12/08/2008 : N° 19154 if FInfo.LoadCompte(FListe.Cells[GCGen,FListe.RowCount-1])
 if FInfo.LoadCompte(FListe.Cells[GCGen,FListe.Row])
    and ( FInfo.GetString('G_VENTILABLE') <> 'X' )
    and ( FListe.Objects[0,FListe.Row] <> nil ) then
  begin
   TVentGuide(FListe.Objects[0,FListe.Row]).Free ;
   FListe.Objects[0,FListe.Row] := nil ;
   exit ;
  end ;

 for iaxe := 1 to MaxAxe do
  begin
   for isec := 0 to T.Ventil[iaxe].Count - 1 do
    if Trim(T.Ventil[iaxe].Strings[isec]) <> '' then
     begin
      lBoExistV := true ;
      break ;
     end ;
  end ; // for

 if lBoExistV then
  if FMsgBox.Execute(33,'','') <> mrYes then
   for iaxe := 1 to MaxAxe do
    for isec := 0 to T.Ventil[iaxe].Count-1 do
     T.Ventil[iaxe].Clear ;

end;

procedure TOM_GUIDE.FDeviseChange( Sender : TObject);
begin
 FInfo.Devise.Load([FDevise.Value]) ;
end;

procedure TOM_GUIDE.SetFoc (C,R : Integer) ;
begin
 FListe.Col:=C ; FListe.Row:=R ; FListe.SetFocus ;
end ;

function TOM_GUIDE.CtrlNumeric ( St : string ; Gene : boolean ): boolean ;
begin
 result := IsMontant(St,Gene) or ( Pos('[',St)>0 ) ;
end ;

function TOM_Guide.CtrlMontant ( Lig :integer) : boolean;
var
 Deb, Cred               : String ;
 CBDeb,CBCred            : Boolean ;
BEGIN
CtrlMontant:=False;
Deb:=FListe.Cells[GCDeb,Lig] ; Cred:=FListe.Cells[GCCre,Lig] ;
CBDeb:=G_Croix(FListe.Cells[GCXDeb,Lig]) ;
CBCred:=G_Croix(FListe.Cells[GCXCre,Lig]) ;

//// ====== CONTROLES POUR TOUS TYPES DE GUIDE ====== ////
if (Deb<>'') and (Cred<>'') then
    BEGIN SetFoc(GCDeb,Lig); FMsgBox.Execute(7,'',''); Exit; END;
if ((DEB<>'') and (Not CtrlNumeric(Deb,False))) or
   ((CRED<>'') and (Not CtrlNumeric(Cred,False))) then
    BEGIN SetFoc(GCDeb,Lig); FMsgBox.Execute(10,'',''); Exit; END;
if CtrlNumeric(Deb,False) and (Pos('[',Deb)<=0) and (Pos('-',Deb)>0) and Not VH^.MontantNegatif then
    BEGIN SetFoc(GCDeb,Lig); FMsgBox.Execute(11,'',''); Exit; END;
if CtrlNumeric(Cred,False) and (Pos('[',Cred)<=0) and (Pos('-',Cred)>0)and Not VH^.MontantNegatif then
    BEGIN SetFoc(GCCre,Lig); FMsgBox.Execute(11,'',''); Exit; END;
if ((Valeur(DEB)<0) and (Pos('[',Deb)<=0) and (Not VH^.MontantNegatif)) then
    BEGIN SetFoc(GCDeb,Lig); FMsgBox.Execute(38,'',''); Exit; END;
if ((Valeur(CRED)<0) and (Pos('[',Cred)<=0) and (Not VH^.MontantNegatif)) then
    BEGIN SetFoc(GCCre,Lig); FMsgBox.Execute(38,'',''); Exit; END;
//// ====== CONTROLES EN FONCTION DU TYPE DE GUIDE ====== ////
if FTypeGuide='ABO' then
   BEGIN
   if ((Deb='') and (Cred='')) or ((Valeur(DEB)=0) and (Valeur(CRED)=0)) then
       BEGIN SetFoc(GCDeb,Lig); FMsgBox.Execute(26,'',''); Exit; END;
   if ((Deb<>'') and (Not IsMontant(Deb,False))) then
       BEGIN SetFoc(GCDeb,Lig); FMsgBox.Execute(26,'',''); Exit; END;
   if ((Cred<>'') and (Not IsMontant(Cred,False))) then
       BEGIN SetFoc(GCCre,Lig); FMsgBox.Execute(26,'',''); Exit; END;
   END else
if ((FTypeGuide='ENC') or (FTypeGuide='DEC')) then
   BEGIN
   // les montants peuvent être non renseignés et sans arrêt
   END else
if ((FTypeGuide='POI') or (FTypeGuide='NOR')) then
   BEGIN
   if (Deb='') and (Cred='') and Not CBDeb and Not CBCred then
       BEGIN SetFoc(GCDeb,Lig); FMsgBox.Execute(17,'',''); Exit; END;
   END;

CtrlMontant:=True;
END;


function TOM_GUIDE.RecupRegle ( Gene,Auxi : String ) : String ;
begin

 if FInfo.LoadAux(Auxi) then
  result := FInfo.GetString('T_MODEREGLE')
   else
    if FInfo.LoadCompte(Gene) and ( ( FInfo.GetString('G_NATUREGENE') = 'TID'  ) or ( FInfo.GetString('G_NATUREGENE') = 'TIC'  ) ) then
     result := FInfo.GetString('G_MODEREGLE')
      else
       result := '' ;

END ;

function TOM_GUIDE.CtrlAuxiGene ( ARow : integer ) : boolean;
var
 CBG          : boolean ;
 CBA          : boolean ;
 lStNatGene   : string ;
 lStNatAuxi   : string ;
 lStNatPiece  : string ;

 procedure _ExitOnError( ACol,Error : integer ) ;
  begin
   SetFoc(Acol,ARow) ;
   FMsgBox.Execute(Error,'','') ;
   Abort ;
  end ;

begin

 result := false ;

 FInfo.Load(FListe.Cells[GCGen,ARow] , FListe.Cells[GCAux,ARow] , GetField('GU_JOURNAL') ) ;

 CBG        := G_Croix(FListe.Cells[GCXGen,ARow]) ;
 CBA        := G_Croix(FListe.Cells[GCXAux,ARow]) ;
 lStNatGene := '' ;
 lStNatAuxi := '' ;

 if ( FInfo.Compte.InIndex > -1 ) then
  begin

   {JP 19/06/06 : Dans les guides d'abonnements, il faut des contrôles plus stricts,
                  notamment avec des comptes entiers}
   if (not CBG) or (FTypeGuide = 'ABO') then
    FListe.Cells[GCGen,ARow] := FInfo.StCompte ;

   lStNatGene := FInfo.GetString('G_NATUREGENE') ;

   if FInfo.GetString('G_FERME') = 'X' then
    _ExitOnError(GCGen,0) ;

   if EstInterdit( FInfo.Journal.GetValue('J_COMPTEINTERDIT') , FInfo.StCompte , 0 ) > 0 then
    begin
     SetFoc(GCGen,ARow) ;
     PgiBox('Ce compte général est interdit sur ce journal.', Ecran.Caption);
     Exit;
   end ;

  end ; // if


 lStNatPiece := GetField('GU_NATUREPIECE') ;

 if ( FInfo.Aux.InIndex > -1 ) then
  begin

   if not CBA then
    FListe.Cells[GCAux,ARow] := FInfo.StAux ;
   {JP 19/06/06 : Ajout de cette qui semble avoir été oubliée}
   lStNatAuxi := FInfo.GetString('T_NATUREAUXI');

   if ( FInfo.GetString('T_FERME') = 'X' ) then
    _ExitOnError(GCAux,1) ;

  end ;

 if VH^.OuiTvaEnc then
  begin
   if (( lStNatGene = '' ) or ( ( lStNatGene <> 'CHA' ) and ( lStNatGene <> 'PRO' ) and ( lStNatGene <> 'IMO' ) ) ) then
     begin
      FListe.Cells[GCXEnc,ARow] := '' ;
      FListe.Cells[GCTva,ARow]  := '' ;
     end ;
  end ;

 if lStNatGene<>'' then
  begin
   if ( (lStNatGene = 'COC') and ( (lStNatPiece = 'FF') or (lStNatPiece = 'AF') or (lStNatPiece = 'RF') or ( lStNatPiece='OF') )) then
    _ExitOnError(GCGen,42) ;

   if ( (lStNatGene = 'COF') and ( (lStNatPiece = 'FC') or (lStNatPiece='AC') or (lStNatPiece='RC') or (lStNatPiece='OC'))) then
    _ExitOnError(GCGen,42) ;
  end ;

 if ((lStNatGene <> '') and (lStNatAuxi <> '')) then
  begin
   if ((lstNatGene = 'COC') and ( (lStNatAuxi = 'AUC') or ( lStNatAuxi = 'FOU') or ( lStNatAuxi='SAL') )) then
    _ExitOnError(GCAux,43) ;

   if ((lStNatGene = 'COF') and ((lStNatAuxi = 'AUD') or (lStNatAuxi='CLI') or (lStNatAuxi='SAL'))) then
    _ExitOnError(GCaux,43) ;

   if ((lStNatGene='COS') and (lStNatAuxi<>'SAL') and (lStNatAuxi<>'DIV')) then
     _ExitOnError(GCaux,43) ;
   {JP 19/06/05 : Ajout de ce teste qui me semble minimum : en abonnement il n'y a pas les tests qui
                  sont effectuées dans les saisies et l'on se retrouve avec des écritures incohérentes}
   if ((lStNatGene <> 'COS') and (lStNatGene <> 'COC') and (lStNatGene <> 'COF') and (lStNatGene <> 'COD'))
      and (lStNatAuxi <> '') then
     _ExitOnError(GCaux,43) ; 
  end ;

 if ( (FTypeGuide <> 'ENC') and (FTypeGuide <> 'DEC') ) and ( FListe.Cells[GCMrg,ARow] = '' ) then
  FListe.CellValues[GCMrg,ARow] := RecupRegle(FListe.Cells[GCGen,ARow],FListe.Cells[GCAux,ARow]) ;

 if ((FTypeGuide='ENC') or (FTypeGuide='DEC')) then
  begin
   if ( ARow = 1 ) and ( FInfo.Compte.InIndex= - 1 ) then
    _ExitOnError(GCGen,27)
  end
   else
    if FTypeGuide = 'ABO' then
     begin
      if ( FInfo.Compte.InIndex= - 1 ) then
       _ExitOnError(GCGen,27) ;
      if ( ( FInfo.GetString('G_COLLECTIF') = 'X') and  ( FInfo.Aux.InIndex = - 1 ) ) then
       _ExitOnError(GCAux,27) ;
    end
     else
      if ( (FTypeGuide='POI') or (FTypeGuide='NOR') ) then
       begin
         if ( FInfo.Compte.InIndex > - 1 ) and ( FInfo.GetString('G_COLLECTIF') = 'X') and ( FListe.Cells[GCAux,ARow] = '' ) and not CBA then
          FListe.Cells[GCXAux,ARow] := 'X' ;
         if ( FListe.Cells[GCGen,ARow] <> '' ) then
          if ( (Pos('*', FListe.Cells[GCGen,ARow]) > 0 ) or ( Pos('?',FListe.Cells[GCGen,ARow]) > 0) or ( FInfo.Compte.InIndex = - 1 ) ) and not CBG then
           _ExitOnError(GCGen,2) ;
         if ( FListe.Cells[GCAux,ARow] <> '' ) then
           if ( (Pos('*',FListe.Cells[GCAux,ARow] ) > 0) or ( Pos('?',FListe.Cells[GCAux,ARow]) > 0 ) or ( FInfo.Aux.InIndex = - 1 ) ) and not CBA then
            _ExitOnError(GCAux,3) ;
         if ( FListe.Cells[GCGen,ARow] = '' ) and ( FListe.Cells[GCAux,ARow] = '' ) and not CBG and not CBA then
          _ExitOnError(GCGen,4) ;
        if ( FListe.Cells[GCAux,ARow] <> '' ) or CBA then
         if ( FInfo.Compte.InIndex > - 1 ) and ( FInfo.GetString('G_COLLECTIF') = '-' ) then
          _ExitOnError(GCGen,5) ;
        if ((FListe.Cells[GCGen,ARow] <> '') and (not CtrlNumeric(FListe.Cells[GCGen,ARow],True))) then
         _ExitOnError(GCGen,9) ;
       end ;

  result := true ;     

end ;


{***********A.G.L.***********************************************
Auteur  ...... : LG
Créé le ...... : 23/04/2007
Modifié le ... : 23/04/2007
Description .. : - LG - 23/04/2007 - FB 17807 - le test du journal en 
Suite ........ : devise etait fait a l'envers
Mots clefs ... : 
*****************************************************************}
function TOM_GUIDE.CtrlEnTete : boolean;
begin

 result := false ;

 if not NomGuideOk then exit ;

 if not FInfo.LoadJournal(GetField('GU_JOURNAL')) then
  begin
   FJournal.SetFocus ;
   FMsgBox.Execute(29,'','') ;
   Exit ;
  end ;

 if FJournal.Value <> '' then
  if ((FJournal.Value = VH^.JalATP) or (FJournal.Value=VH^.JalVTP)) then
   begin
    FMsgBox.Execute(46,'','') ;
    FJournal.SetFocus ;
    Exit ;
   end ;

// if FNaturePiece.Text='' then
  {JP 14/06/06 : FQ 18302 : je me demande si le tous est bien utile !!
                 je laisse FNaturePiece.Text car Laurent m'a dit que c'était ainsi dans Guide.Pas
                 et qu'il faut laisser le Tous ...}
  if (FNaturePiece.Text = '') or ((FNaturePiece.Value = '') and (FTypeGuide = 'ABO')) then
  begin
    FNaturePiece.SetFocus ;
    FMsgBox.Execute(30,'','') ;
    Exit ;
  end;

  {JP 14/06/06 : FQ 18302}
  if (VarToStr(GetField('GU_ETABLISSEMENT')) = '') and (FTypeGuide = 'ABO') then begin
    SetFocusControl('GU_ETABLISSEMENT');
    LastError := 1;
    LastErrorMsg := TraduireMemoire('Pour les guides d''abonnements, l''établissement est obligatoire.');
  end;

 if FDevise.Value = '' then
  begin
   FMsgBox.Execute(35,'','');
   FDevise.SetFocus ;
   Exit ;
  end ;


 if ( FDevise.value <> V_PGI.DevisePivot ) then
  begin

   if FInfo.Journal.GetValue('J_MULTIDEVISE') = '-' then
    begin
     FJournal.SetFocus ;
     FMsgBox.Execute(14,'','') ;
     Exit ;
    end ; // if

  end;

 result := true ;

end ;


function TOM_GUIDE.CtrlCompta : boolean ;
var
 C, D : double ;
 i : integer ;
begin

 result := false ;

 if FTypeGuide = 'ABO' then
  begin

   C := 0 ;
   D := 0 ;

   for i := 1 to FListe.RowCount - 1 do
    begin
     C := C + Valeur(FListe.Cells[GCCre,i]) ;
     D := D + Valeur(FListe.Cells[GCDeb,i]) ;
    end; // for

   if ( Arrondi( D-C ,V_PGI.OkDecE ) <> 0 ) then
    begin
     FMsgBox.Execute(32,'','') ;
     exit;
    end ;

  end ;

  if ( FInfo.Journal.GetValue('J_NATUREJAL') = 'BQE' ) or ( FInfo.Journal.GetValue('J_NATUREJAL') = 'CAI' ) then
   begin

    for i := 1 to FListe.RowCount - 1 do
     begin
      result := Fliste.Cells[GCGen,i] = FInfo.Journal.GetValue('J_CONTREPARTIE') ;
      if result then break ;
     end ; // for

   if not result then
    begin
     FMsgBox.Execute(44,'','') ;
     exit ;
    end ;

   end ;

 result := true ;

end ;



function TOM_GUIDE.LigneAvecArret( ARow : integer ) : boolean ;
begin
 result := ( ( UpperCase(FListe.Cells[GCXGen,ARow]) = 'X' ) or ( UpperCase(FListe.Cells[GCXAux,ARow]) = 'X' ) ) and
           ( ( UpperCase(FListe.Cells[GCXDeb,ARow]) = 'X' ) or ( UpperCase(FListe.Cells[GCXCre,ARow]) = 'X' ) ) ;
end ;

function TOM_GUIDE.CtrlGuideLigne ( ARow : integer ) : boolean ;
begin

 result := CtrlMontant(ARow) ;
 if not result then exit ;
 result := CtrlAuxiGene(ARow) ;

end ;


function TOM_GUIDE.CtrlGuide : boolean ;
var
 i : integer ;
 lErr : integer ;
begin

 result := false ;

 if FListe.RowCount = 2 then
  begin
   FMsgBox.Execute(28,'','') ;
   exit ;
  end ; // if

 if G_LigneVide(FListe,FListe.RowCount - 1) and not LigneAvecArret(FListe.RowCount - 1 ) then
  FListeDeleteRow(FListe.RowCount - 1) ; 

 if not CtrlEnTete then exit ;
 if not CtrlCompta then exit ;

 for i := 1  to FListe.RowCount - 1 do
  begin

   if G_LigneVide(FListe,i) and not LigneAvecArret(i) then
    begin
     FMsgBox.Execute(25,'','') ;
     exit ;
    end ;

   if not CtrlGuideLigne(i) then exit ;
   lErr := CtrlFormule(FListe,i,false) ;
   if ( lErr <> - 1 ) then
    begin
     FMsgBox.Execute(lErr,'','') ;
     exit ;
    end ;
   
  end ; // for

 result := true ;

end;

procedure TOM_GUIDE.BVentilClick( Sender : TObject);
var
 T : TVentGuide ;
begin

 if (length(FListe.Cells[GCGen,FListe.Row]) < VH^.Cpta[fbGene].Lg ) or not FInfo.LoadCompte(FListe.Cells[GCGen,FListe.Row]) then
  begin
   SetFoc(GCGen,FListe.Row) ;
   FMsgBox.Execute(16,'','') ;
   exit ;
  end ;

  if DS.State = dsBrowse then
   DS.Edit ;

 if ( FInfo.GetString('G_VENTILABLE') = 'X' ) then
  begin
   T := TVentGuide(FListe.Objects[0,FListe.Row]) ;
   if T = nil then
    begin
     T                            := TVentGuide.Create ;
     FListe.Objects[0,FListe.Row] := T ;
    end ; // if

   ParamGuideAna(T,FListe.Cells[GCGen,FListe.Row],FTypeGuide);

   end
    else
     begin
      SetFoc(GCGen,FListe.Row) ;
      FMsgBox.Execute(16,'','') ;
     end;

 FListe.SetFocus ;

end;

procedure TOM_GUIDE.BEcheClick( Sender : TObject );
begin
 FListe.CellValues[GCMrg,FListe.Row] := FicheRegle_AGL(FListe.CellValues[GCMrg,FListe.Row],true,TFFiche(Ecran).FTypeAction);
 FListe.SetFocus ;
 if DS.State = dsBrowse then
  DS.Edit ;
end;

procedure TOM_GUIDE.BInsLigneClick(Sender: TObject);
begin
 CreateRow(FListe.Row,true) ;
end;

procedure TOM_GUIDE.BDelLigneClick( Sender : TObject);
begin
 FListeDeleteRow( FListe.Row ) ;
end;

procedure TOM_GUIDE.NextRow ;
var
 lBoCancel : boolean ;
 lBoChg    : boolean ;
begin

 lBoChg    := true ;
 lBoCancel := false ;
 FListeRowExit( nil , FListe.Row , lBoCancel, lBoChg );
 if lBoCancel then exit ;
 CreateRow ;
 EnableButtons ;

end;

procedure TOM_GUIDE.CreateRow( ARow : integer = -1 ; vBoInsert : boolean = false );


 procedure _Init ;
  begin
  // InitRow ;
   FListe.Col      := GCGen ;
   if FListe.CanFocus then FListe.SetFocus ;
   FListe.ShowEditor ;
  end;

 procedure _InsererLigne ;
  begin
   FListe.InsertRow(ARow) ;
   FListe.Row      := ARow ;
    G_Renum(FListe) ;
   FListe.Refresh ;
  end;

 procedure _AjouterLigne ;
  begin
   FListe.RowCount := FListe.RowCount + 1 ;
   FListe.Row      := FListe.RowCount - 1 ;          // on se place sur cette nouvelle ligne
   ARow            := FListe.Row ;
  end;

begin

 if ARow = - 1 then ARow := FListe.Row ;

  if vBoInsert then
   _InsererLigne
    else
     begin
      if ( FListe.Row = FListe.RowCount - 1 )  then
       _AjouterLigne
        else
         if ( FListe.Row > 1 ) and ( FListe.Row < FListe.RowCount  ) then
          _InsererLigne ;
     end; // if

  EnableButtons ;

  _Init ;

 if DS.State = dsBrowse then
  DS.Edit ;

end;

procedure TOM_GUIDE.FListeDeleteRow( Row : integer ) ;
begin

 if ( (Fliste.Row <= 0) or (FListe.RowCount <= 2) ) then exit ;

 if Row = - 1 then Row := FListe.Row ;

 if FListe.RowCount = 2 then
  begin
   FListe.VidePile(False) ;
   CreateRow ;
  end
   else
    begin
     if ( Row <> 1 ) and ( FListe.Row = Row ) then // si on n'est pas sur la derniere cellule on remonte d'une ligne
      FListe.Row := FListe.Row - 1 ;
     if FListe.Objects[0,Row] <> nil then
      TVentGuide(FListe.Objects[0,Row]).Free ;
     FListe.Objects[0,Row] := nil ;
     FListe.DeleteRow(Row) ; // on supprime la ligne
     EnableButtons ;
    end; // if

 G_Renum(FListe) ;
 FListe.SetFocus ;

 if DS.State = dsBrowse then
  DS.Edit ;

end;

procedure TOM_GUIDE.FListeDblClick(Sender: TObject);
begin

 case FListe.Col of
   GCGen : ChercheGen ;
   GCAux : ChercheAux ;
   GCXGen,GCXAux,GCXRef,GCXLib,GCXDeb,GCXCre,GCXMrg,GCXEnc :
           begin
           if UpperCase(FListe.Cells[FListe.Col,FListe.Row])='X' then
            FListe.Cells[FListe.Col,FListe.Row]:='-'
              else
               FListe.Cells[FListe.Col,FListe.Row]:='X' ;
           end ;
   else
    AssisteGuide ;
 end ; // if

 if DS.State = dsBrowse then
  DS.Edit ;

end;

procedure TOM_GUIDE.AssisteGuide ;
begin

 if ( FListe.Col Mod 2 ) <> 0 then exit ;

 if ( ( FListe.Col <= GCAux ) or ( FListe.Col >= GCTva ) ) then exit ;

 FListe.Cells[FListe.Col,FListe.Row] := FListe.Cells[FListe.Col,FListe.Row] + ChoixChampZone(FListe.Row,'GUI') ;
 FListe.SetFocus ;

 if DS.State = dsBrowse then
  DS.Edit ;

end;


{***********A.G.L.***********************************************
Auteur  ...... : LG
Créé le ...... : 25/04/2006
Modifié le ... : 17/05/2006
Description .. : - LG - 25/04/2006 - FB 17930 - alimentation du champs
Suite ........ : gu_datecreation
Suite ........ : - LG - 17/05/2006- FB 18170 - la duplication de fct aps, on
Suite ........ : enregistre le guide pour passer a l'etta dsBrowe
Mots clefs ... :
*****************************************************************}
procedure TOM_GUIDE.BDupliquerClick (Sender : TObject) ;
var
 lStJournal      : string ;
 lStDevise       : string ;
 lStNature       : string ;
 lStEta          : string ;
 lBoTreso   : boolean ;
begin

 if DS.State <> dsBrowse then
  begin
   TFFiche(Ecran).BValiderClick(nil) ;
   if LastError <> 0 then exit ;
  end ;

 lStJournal           := GetField('GU_JOURNAL') ;
 lStDevise            := GetField('GU_DEVISE') ;
 lStNature            := GetField('GU_NATUREPIECE') ;
 lStEta               := GetField('GU_ETABLISSEMENT') ;
 lBoTreso   := GetCheckBoxState('GU_TRESORERIE') = cbChecked ;

 DS.Insert ;
 OnNewRecord ;

 SetField('GU_JOURNAL'        , lStJournal ) ;
 SetField('GU_DEVISE'         , lStDevise ) ;
 SetField('GU_NATUREPIECE'    , lStNature ) ;

 if not VH^.ProfilUserC[prEtablissement].ForceEtab and VH^.EtablisCpta then
   {JP 16/10/07 : FQ 16149 : Gestion des restrictions établissement : on force au besoin l'établissement}
   SetField('GU_ETABLISSEMENT'  , lStEta ) ;

 {b Thl 13/06/2006 FQ 18068}
 // SetField('GU_TRESORERIE'     , lBoTreso ) ;

 if lBoTreso then
   SetField('GU_TRESORERIE', 'X')
 else
   SetField('GU_TRESORERIE', '-');

 {e Thl 13/06/2006}
 SetField('GU_DATECREATION'   , Now ) ;
end;


{***********A.G.L.***********************************************
Auteur  ...... : LG
Créé le ...... : 25/04/2006
Modifié le ... :   /  /    
Description .. : - LG  - 25/04/2006 - FB 17930 - on reprends la date de ref 
Suite ........ : si'l definie pour tester les guide en mode bordereau
Mots clefs ... : 
*****************************************************************}
procedure TOM_GUIDE.BGuideClick( Sender : TObject);
{$IFDEF COMPTA}
var
 M : RMVT ;
 DateCpt : TDateTime ;
{$ENDIF}
begin
{$IFDEF COMPTA}
 if not TToolbarButton97(GetControl('BGuide',true)).Enabled then Exit;
 TFFiche(Ecran).BValiderClick(nil) ;
 if LastError <> 0 then exit ;
 Application.ProcessMessages ;

 FillChar(M,Sizeof(M),#0) ;
 M.TypeGuide:=FTypeGuide ; M.LeGuide:=FQuelGuide ; M.FromGuide:=True ;
 M.Simul:='S' ; M.Etabl:=GetField('GU_ETABLISSEMENT') ;
 M.Jal:=GetField('GU_JOURNAL') ; M.DateC:=V_PGI.DateEntree ;
 if VH^.CPExoRef.Code <> '' then
  M.DateC := VH^.CPExoRef.Fin
   else
    M.DateC:=V_PGI.DateEntree ;
 M.DateTaux:=M.DateC ; DateCpt:=M.DateC ;
 M.Exo:=QuelExoDT(M.DateC) ;
 M.CodeD:=GetField('GU_DEVISE') ; M.TauxD:=GetTaux(M.CodeD,DateCpt,M.DateC) ;
 M.Nature:=GetField('GU_NATUREPIECE') ;

 FInfo.Journal.Load([GetField('GU_JOURNAL')] ) ;

 if ( FInfo.Journal.GetValue('J_MODESAISIE') = '' ) or
    ( FInfo.Journal.GetValue('J_MODESAISIE') = '-' ) then
  LanceSaisie(Nil,taCreat,M) 
   else
    LanceSaisieFolioGuide(M) ;

 FListe.SetFocus ;

{$ENDIF}
end;

function TOM_GUIDE.NomGuideOk : boolean ;
begin

 result := true ;

 if GetField('GU_LIBELLE') = '' then
  begin
   FMsgBox.Execute(20,'','') ;
   Result:=false ;
   Exit ;
  end ;

 if DS.State <> dsInsert then
  exit ;

 if ExisteSQL('Select GU_GUIDE From GUIDE Where GU_TYPE="' + FTypeGuide + '" AND GU_LIBELLE="' + GetField('GU_LIBELLE') + '"' ) then
  begin
   FMsgBox.Execute(41,'','') ;
   SetFocusControl('GU_LIBELLE') ;
   Result := False ;
   exit ;
  end ;

end ;

procedure TOM_GUIDE.FJournalChange(Sender: TObject);
begin
 (* JP 19/06/06
 if FTypeGuide <> 'ABO' then
  FDEVISE.Enabled := true    dans le OnArgument
   else
    begin
     FDEVISE.Value    := V_PGI.DevisePivot ; fait dans le onNewRecord donc inutile ici
     FDEVISE.Enabled  := False ; dans le OnArgument
    end ;
 *)

 if FInfo.LoadJournal(FJournal.Value) then
  begin
   FNaturePiece.DataType := FInfo.Journal.GetTabletteNature ;
   if DS.State in [dsInsert, dsEdit] then begin
     FNaturePiece.Value    := FInFo.Journal.getValue('J_NATDEFAUT') ;
     if FNaturePiece.Value = '' then
      FNaturePiece.ItemIndex := 0 ;
   end;
  end ;

end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/03/2006
Modifié le ... :   /  /
Description .. : 28/03/2006 - FB 17656 - ajout du bouton assistant
Mots clefs ... :
*****************************************************************}
procedure TOM_GUIDE.BAssistantClick(Sender: TObject);
begin
 AssisteGuide ;
 FListe.SetFocus ;
end;

{JP 15/10/07 : FQ 16149 : gestion des réstrictions Etablissements et à défaut des ParamSoc
{---------------------------------------------------------------------------------------}
procedure TOM_GUIDE.GereEtablissement;
{---------------------------------------------------------------------------------------}
begin
  {Si l'on ne gère pas les établissement ...}
  if not VH^.EtablisCpta  then begin
    {... on affiche l'établissement par défaut}
    SetControlText('GU_ETABLISSEMENT', VH^.EtablisDefaut);
    SetField('GU_ETABLISSEMENT', VH^.EtablisDefaut);
    {... on désactive la zone}
    SetControlEnabled('GU_ETABLISSEMENT', False);
  end

  {On gère l'établissement, donc ...}
  else begin
    {... On commence par regarder les restrictions utilisateur}
    PositionneEtabUser(GetControl('GU_ETABLISSEMENT'));
    {... s'il n'y a pas de restrictions, on reprend le paramSoc
     JP 25/10/07 : FQ 19970 : Finalement on oublie l'option de l'établissement par défaut
    if GetControlText('GU_ETABLISSEMENT') = '' then begin
      {... on affiche l'établissement par défaut
      SetControlText('GU_ETABLISSEMENT', VH^.EtablisDefaut);
      SetField('GU_ETABLISSEMENT', VH^.EtablisDefaut);
      {... on active la zone
      SetControlEnabled('GU_ETABLISSEMENT', True);
    end;}
  end;
end;




Initialization
  registerclasses ( [ TOM_GUIDE ] ) ;
end.

