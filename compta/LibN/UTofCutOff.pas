{$A+,B-,C+,D+,E-,F-,G+,H+,I+,J+,K-,L+,M-,N+,O+,P+,Q-,R-,S-,T-,U-,V+,W-,X+,Y+,Z1}
{$MINSTACKSIZE $00004000}
{$MAXSTACKSIZE $00100000}
{$IMAGEBASE $00400000}
{$APPTYPE GUI}
Unit UTofCutOff ;

Interface

Uses
     Classes,
     UTof,
     UTob,
     HCtrls,
     Menus,           // PopUpMenu
     Grids,           // TGridDrawState
     Graphics,
{$IFDEF EAGLCLIENT}
     MainEAgl,        // AGLLanceFiche
     UtileAgl, {LanceEtatTob}
{$ELSE}
     Fe_Main,         // AGLLanceFiche
     Db,
     EdtREtat, {LanceEtatTob}
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     ULibEcriture ,
     uTofViergeMul;   // Fiche Ancetre Vierge MUL

Type



 TCutOff = Class(TObjetCompta)
  private
   FTOBN1       : TOB ;
   FTOBN3       : TOB ;
   FTOB         : TOB ;
   FDateC       : TDateTime ;
   FJournal     : string ;
   FQualifPiece : string ;
   FCR          : TList ;
   FTOBGene     : TOB ;
  function    ChargeInfo : boolean ;
  function    MAJEcriture : boolean ;
  procedure   AjouteN3aN1 ;
  procedure   Generer ;
  function    Construire : boolean ;
  procedure   StockeLesEcrituresIntegrees( vTOB : TOB) ;
  procedure   TransfertAna ( vTOBEcr,vTOBCut : TOB ) ;
  function    GetJournal ( vTOB : TOB ) : string ;
  function    GetQualifPiece : string ;
  function    MakeCrit( vTOB : TOB ) : string ;
  public
   constructor Create( vInfoEcr : TInfoEcriture ) ; override ;
   destructor  Destroy ; override ;
   function    Execute( vTOB : TOB ; vDateC : TDateTime ; vJournal,vQualifPiece : string ) : boolean ;
   procedure   RAZFLC ;
   property    CR    : TList read FCR ;
  end;


  TOF_CPMULCUTOFF = Class (TOF_ViergeMul)
    private
     FStCompte            : string ;
     FTCutOff             : TCutOff ;
     FInfo                : TInfoEcriture ;
     FMessCompta          : TMessageCompta ;        // affichage des messages
     PopUpTraitement      : TPopUpMenu;
     PopUpUtilitaire      : TPopUpMenu;
     PopUpEtat            : TPopUpMenu;
     E_JOURNAL            : THValComboBox ;
    procedure OnExitE_General         ( Sender : TObject );
    procedure OnExitDateC             ( Sender : TObject );
    {$IFNDEF EAGLCLIENT}
    procedure SaisieEcrClick          ( Sender : TObject ); // Saisie d'une Piece
    {$ENDIF}
    procedure SaisieBorClick          ( Sender : TObject ); // Saisie d'un Bordereau
    procedure OnClickInfosComp        ( Sender : TObject ); // Informations complementaires de l'ecriture
    procedure OnClickCompteGeneral    ( Sender : TObject ); // Paramètres du compte général
    procedure OnElipsisClickE_General ( Sender : TObject );
    procedure GetCellCanvasFListe     ( ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState) ;
    procedure FListePostDrawCell      ( ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState ) ;
    procedure QUALIFPIECEChange       ( Sender : TObject ) ;
    procedure JOURNALChange           ( Sender : TObject ) ;
    procedure OnClickBGENE            (Sender: TObject);
    procedure OnDblClickFListe        ( Sender : TObject);
    procedure OnPopUpPopUpTraitement  ( Sender : TObject );
    procedure OnPopUpPopUpUtilitaire  ( Sender : TObject );
    procedure OnPopUpEtat             ( Sender : TObject );
    procedure InitPopUp               ( vActivation : Boolean );
//    procedure OnClickSansPresta       ( Sender : TObject );
    procedure OnClickEtatJustifCompte ( Sender : TObject);
    procedure OnErrorTOB              ( sender : TObject ; Error : TRecError);
   protected
    procedure RemplitATobFListe ; override ;
    procedure InitControl ; override ;
    function  AjouteATobFListe( vTob : Tob) : Boolean; override ;
    function  BeforeLoad : boolean ; override ;
   public
    procedure OnLoad                   ; override ;
    procedure OnArgument              (S : String ) ; override ;
    procedure OnClose ; override ;
    procedure AfterShow;               override;
    procedure OnKeyDownEcran          ( Sender : TObject; var Key : Word; Shift : TShiftState); override;
    procedure OnPopUpPopF11           ( Sender : TObject ); override ;
    procedure OnClickBImprimer  (Sender : Tobject); override ;
  end ;


  TOF_CPCUTOFFPARAM = Class (TOF)
    procedure BValiderClick(Sender: TObject); // Bouton Valider
   public
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure OnLoad                   ; override ;
  end ;



function CPLanceFiche_CPMULCUTOFF( vStParam : string = '' ) : string ;

Implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  CPTypeCons,
  {$ENDIF MODENT1}
  ULibAnalytique,
  ed_tools,
  ImRapInt,
  LookUp,        // LookUpList
  SaisUtil,      // RMVT
  Ent1 ,
  SaisComm ,
  uLibWindows ,
  Vierge ,
  PARAMSOC,        // GetParamSocSecur
  Windows,         // VK_
  Htb97,           // TToolBarButton97
  HQry,            // HQuery
  HEnt1,
  HMsgBox,
  Sysutils,
  StdCtrls,
  Controls,
  Saisie,           // TrouveSaisie
  CPTiers_TOM,      // FicheTiers
  SaisComp,         // R_COMP
  UTOFCPMULMVT,
  CPGeneraux_TOM,   // FicheGene
  UlibExercice ,
  Constantes ,
  SaisBor ;          // LanceSaisieFolio


const
 cFI_TABLE    = 'CPMULCUTOFF' ;


function _RecupEcrSelect ( FListe : THGrid ; vTOB : TOB = nil ) : TOB ;
var
 lQuery   : Tquery ;
 lSQL     : string;
begin

  result := Tob.Create('ECRITURE', nil, -1);

  if vTOB = nil then
   vTOB := GetO(FListe, FListe.Row);

  lSQL := 'SELECT * FROM ECRITURE WHERE ' +
         'E_JOURNAL= "' + vTOB.GetValue('E_JOURNAL') + '" AND ' +
         'E_EXERCICE = "' + vTOB.GetValue('E_EXERCICE') + '" AND ' +
         'E_DATECOMPTABLE = "' + USDateTime(vTOB.GetValue('E_DATECOMPTABLE')) + '" AND ' +
         'E_NUMEROPIECE = ' + IntToStr(vTOB.GetValue('E_NUMEROPIECE')) + ' AND ' +
         'E_NUMLIGNE = ' + IntToStr(vTOB.GetValue('E_NUMLIGNE')) + ' AND ' +
         'E_NUMECHE = ' + IntToStr(vTOB.GetValue('E_NUMECHE')) + ' AND ' +
         'E_QUALIFPIECE = "' + vTOB.GetValue('E_QUALIFPIECE') + '"';

 // JP 01/08/05 : FQ 15124 : On exclut les écritures TTC de la Gescom}
 lSQL  := lSQL + ' AND (E_REFGESCOM = "" OR E_REFGESCOM IS NULL OR E_TYPEMVT <> "TTC")' ;
 //JP 26/06/07 : FQ TRESO 10491 : on vérouille les flux originaires de la Tréso 
 lSQL  := lSQL + ' AND (E_QUALIFORIGINE <> "' + QUALIFTRESO + '" OR E_QUALIFORIGINE IS NULL OR E_QUALIFORIGINE = "")' ;

 lQuery := OpenSQL(lSQL, True);
  try
    if not lQuery.Eof then
      Result.SelectDB('', lQuery );

  finally
    Ferme( lQuery );
  end;
end;


function CPLanceFiche_CPMULCUTOFF( vStParam : string = '' ) : string ;
var
 lInfo : TInfoEcriture ;
begin

 lInfo := TInfoEcriture.Create ;
 
 try

 if not lInfo.LoadCompte( GetParamSocSecur('SO_CPCCA','')) or not lInfo.LoadCompte( GetParamSocSecur('SO_CPPCA','')) then
  begin
   PGIInfo('Les comptes de charge périodiques ne sont pas renseignés.','Gestion des charges périodiques');
   exit ;
  end ;

 finally
  lInfo.Free ;
 end ;

  Result := AGLLanceFiche('CP', 'CPMULCUTOFF', '', '', vStParam);
  
end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 12/09/2005
Modifié le ... :   /  /    
Description .. : - FB 15890 - 12/09/2005 - corection du calcul pour els 
Suite ........ : compte de charges
Mots clefs ... : 
*****************************************************************}
function TOF_CPMULCUTOFF.AjouteATobFListe( vTob : Tob): Boolean ;
var
 vD1,vD2  : double ;
 lRdSolde : double ;
begin
 lRdSolde := vTob.GetValue('E_DEBIT') - vTob.GetValue('E_CREDIT') ;
 vTob.AddChampSupValeur('SOLDE', lRdSolde ) ;
 if vTob.GetValue('EC_CUTOFFDATECALC') <> iDate1900 then
  vTob.AddChampSupValeur('BOGENEREE', 'X' )
   else
    vTob.AddChampSupValeur('BOGENEREE', '-' ) ;
 vD2 := vTob.GetValue('EC_CUTOFFFIN') ;
 vD1 := vTob.GetValue('EC_CUTOFFDEB') ;
 vTob.AddChampSupValeur('SELECTED', '-');
 vTob.AddChampSupValeur('DUREE', vD2 - vD1 + 1 ) ;
 vD1 := StrToDate(GetControlText('DATEC')) + 1 ;
 vTob.AddChampSupValeur('BASE', vD2 - vD1 + 1 ) ;
 if vTob.GetValue('BASE') < 0 then
  vTob.PutValue('BASE', 0) ;
 if vTob.GetValue('BASE') > vTob.GetValue('DUREE') then
  vTob.PutValue('BASE', vTob.GetValue('DUREE') ) ;
 vTob.AddChampSupValeur('MCCA', 0 ) ;
 vTob.AddChampSupValeur('MPCA', 0 ) ;
 if vTob.GetValue('G_NATUREGENE') = 'CHA' then
  begin
   vTob.PutValue('MCCA' , Arrondi(( lRdSolde / vTob.GetValue('DUREE') ) * vTob.GetValue('BASE'),0) ) ;
   if ( vTob.GetValue('MCCA') < 0 ) and not VH^.MontantNegatif then
    begin
     vTob.PutValue('MPCA' , vTob.GetValue('MCCA')*(-1)) ;
     vTob.PutValue('MCCA' , 0 ) ;
    end ;
  end
    else
     begin
      vTob.PutValue('MPCA' , Arrondi(( ( lRdSolde * (-1) ) / vTob.GetValue('DUREE') ) * vTob.GetValue('BASE'),0) ) ;
      if ( vTob.GetValue('MPCA') < 0 ) and not VH^.MontantNegatif then
       begin
        vTob.PutValue('MCCA' , vTob.GetValue('MPCA')*(-1)) ;
        vTob.PutValue('MPCA' , 0 ) ;
       end ;
     end ;
 {$IFDEF TT}
 // vTob.AddChampSupValeur('SOLDE', lRdVal ) ;
 {$ENDIF}
 result := true ;
end;

procedure TOF_CPMULCUTOFF.InitControl;
var
 lDtDateC : TDateTime ;
begin

 PageControl.ActivePage := PageControl.Pages[0];

 if VH^.CPExoRef.Code<>'' then
  lDtDateC := VH^.CPExoRef.Fin
   else
    lDtDateC := VH^.Encours.Fin ;

 SetControlText('DATEC',DateToStr(lDtDateC) ) ;

 THEdit(GetControl('G_GENERAL', True)).MaxLength    := VH^.CPta[fbGene].Lg ;
 THEdit(GetControl('G_GENERAL_', True)).MaxLength   := VH^.CPta[fbGene].Lg ;
 SetControlProperty('ECRGEN', 'State', cbGrayed ) ;
 SetControlChecked('EXHAUS' , false ) ;
 SetControlChecked('SSPRESTA' , false ) ;
 SetControlText('QUALIFPIECE','N') ;
// SetControlText('JOURNAL' , GetParamSocSecur('SO_CPCUTJAL' , '' ) ) ;


end;

procedure TOF_CPMULCUTOFF.AfterShow ;
begin
 inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 19/05/2005
Modifié le ... : 13/10/2005
Description .. : - FB 15884 - 19/05/2005 - on ne peut pas rentrer tant que 
Suite ........ : l'on a pas renseigne les param societe
Suite ........ : - FB 15890 - 12/09/2005 - ajout des journaux de regul ds la 
Suite ........ : tablette des,journaux dispo
Suite ........ : - FB 16870 - 13/10/2005 - on choisir l premeir journal d'OD 
Suite ........ : ds la liste
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPMULCUTOFF.OnArgument (S : String ) ;
var
 i         : integer ;
 lStJalDef : string ;
begin

 {$IFDEF TT}
 V_PGI.SAV                           := true ;
 {$ENDIF}

 Ecran.HelpContext := 7248000 ;

 FFI_Table                           := cFI_TABLE;
 FStListeParam                       := 'CPMULCUTOFF';
 FStCompte                           := S ;
 inherited ;

 //BImprimer.OnClick                   := nil ;

 PopUpTraitement                     := TPopUpMenu(GetControl('PopUpTraitement', True));
 PopUpUtilitaire                     := TPopUpMenu(GetControl('PopUpUtilitaire', True));
 PopUpEtat                           := TPopUpMenu(GetControl('POPETAT', True));
 E_JOURNAL                           := THValComboBox(GetControl('JOURNAL', True)) ;
 E_JOURNAL.DataType                  := 'TTJALCRIT' ;
 E_JOURNAL.Plus                      := ' and ( J_NATUREJAL="REG" or J_NATUREJAL="OD" OR J_NATUREJAL="EXT"  ) ' +
                                        ' or ( ( J_MODESAISIE="BOR" or J_MODESAISIE="LIB" ) and  (J_NATUREJAL="REG" OR J_NATUREJAL="OD" OR J_NATUREJAL="EXT" ) ) ';
 E_JOURNAL.ReLoad ;
 FListe.GetCellCanvas                := GetCellCanvasFListe ;
 FListe.OnDblClick                   := OnDblClickFListe ;
 PopUpTraitement.OnPopUp             := OnPopUpPopUpTraitement ;
 PopUpUtilitaire.OnPopUp             := OnPopUpPopUpUtilitaire;
 PopUpEtat.OnPopUp                   := OnPopUpEtat;
 PopF11.OnPopup                      := OnPopUpPopF11 ;
 Ecran.OnKeyDown                     := OnKeyDownEcran  ;
 FListe.PostDrawCell                 := FListePostDrawCell;
 FInfo                               := TInfoEcriture.Create ;
 FTCutOff                            := TCutOff.Create(FInfo) ;
 FTCutOff.OnError                    := OnErrorTOB ;
 FMessCompta                         := TMessageCompta.Create('Traitement des charges périodes') ;

 TToolBarButton97(GetControl('BVALIDER', True)).OnClick   := OnClickBGENE;
 THEdit(GetControl('G_GENERAL'  , True)).OnElipsisClick   := OnElipsisClickE_General;
 THEdit(GetControl('G_GENERAL'  , True)).OnDblClick       := OnElipsisClickE_General;
 THEdit(GetControl('G_GENERAL'  , True)).OnExit           := OnExitE_General;
 THEdit(GetControl('G_GENERAL_' , True)).OnElipsisClick   := OnElipsisClickE_General;
 THEdit(GetControl('G_GENERAL_' , True)).OnDblClick       := OnElipsisClickE_General;
 THEdit(GetControl('G_GENERAL_' , True)).OnExit           := OnExitE_General;
// THEdit(GetControl('SSPRESTA'   , True)).OnClick          := OnClickSansPresta ;

 lStJalDef := GetParamSocSecur('SO_CPCUTJAL' , '' ) ;

 if lStJalDef <> '' then
  SetControlText('JOURNAL' , lStJalDef )
  else
   for i := 0 to E_JOURNAL.Items.Count - 1 do
    begin
      if not FInfo.LoadJournal(E_JOURNAL.Values[i]) then continue ;
      if FInfo.Journal.GetValue('J_NATUREJAL') = 'OD' then
       begin
        E_JOURNAL.ItemIndex := i ;
        break ;
       end ;
    end ; // if

 E_JOURNAL.OnChange                                              := JOURNALChange ;
 THValComboBox(GetControl('QUALIFPIECE', True)).OnChange         := QUALIFPIECEChange ;
 THEdit(GetControl('DATEC', True)).OnExit                        := OnExitDateC ;

 JOURNALChange(nil) ;

 //SetControlEnabled('BVALIDER' , false ) ;

 //CUpdateGridListe ( FListe , FStListeParam ) ;

  // GCO - 16/09/2005 - FQ 16678
  if FStCompte <> '' then
  begin
    SetControlText('G_GENERAL', FStCompte ) ;
    SetControlText('G_GENERAL_', FStCompte ) ;
  end ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : LG
Créé le ...... : 14/02/2007
Modifié le ... :   /  /    
Description .. : - FB 18700 - 14/02/2007 - on ne peut pas valider ss ligne 
Suite ........ : selectionné
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPMULCUTOFF.OnKeyDownEcran (Sender: TObject; var Key: Word; Shift: TShiftState);
var
 lTOB     : TOB ;
 lBoBor   : boolean ;
 //lBoActif : boolean ;
begin

 //inherited;
 case Key of
  67 : begin // Alt + C -> Informations complémentaires
        if ssAlt in Shift then
         OnClickInfosComp(nil);
       end;
 VK_SPACE : begin
             lTOB := GetO(FListe, FListe.Row) ;
             if lTOB = nil then exit ;
             lBoBor := ( lTOB.GetValue('E_MODESAISIE') = 'BOR' ) or ( lTOB.GetValue('E_MODESAISIE') = 'LIB' ) ;
             if lBoBor and ( GetControlText('QUALIFPIECE') <> 'N' ) then
              Key := 0
               else
                inherited ;
            // lBoActif := ( FListe.AllSelected or ( FListe.nbSelected > 0 ) ) and ( GetCheckBoxState('SSPRESTA') <> cbChecked ) ;
             //lBoActif := ( GetCheckBoxState('SSPRESTA') <> cbChecked ) ;
             //SetControlEnabled('BVALIDER' , lBoActif ) ;
            end ;
  else
   inherited ;
 end ;// case


end ;


procedure TOF_CPMULCUTOFF.OnClose;
begin
 FreeAndNil(FTCutOff) ;
 FreeAndNil(FInfo) ;
 FreeAndNil(FMessCompta) ;
 inherited;
end;


procedure TOF_CPMULCUTOFF.OnLoad ;
begin
  Inherited ;
  InitAutoSearch ;
  InitPopUp(false) ;
end ;



{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 19/05/2005
Modifié le ... : 19/07/2007
Description .. : - FB 15882 - on tri par general puis date
Suite ........ : - LG - FB 15882 - 12/10/2005 - force l'ordre de tri pur la 
Suite ........ : fiche ancetre
Suite ........ : - LG - FB 19011 - supprime le champs e_datecomptable en 
Suite ........ : double
Suite ........ : - LG - FB 16806 - 
Suite ........ : - LG - FB 19769 - pour les ecritures ss cutoff, on tient 
Suite ........ : compte de l'exercice de l'ecriture
Suite ........ : - LG - FB 19769 - on appliquer l'exercide pour tout les types 
Suite ........ : de calcul
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPMULCUTOFF.RemplitATobFListe;
var
 lStEcrGen      : string ;
 lStExhau       : string ;
 lState         : TCheckBoxState ;
 lStatePresta   : TCheckBoxState ;
 lStSqlCompl    : string ;
 lStExo         : TExoDate ;
begin
 lState        := GetCheckBoxState('ECRGEN') ;
 lStatePresta  := GetCheckBoxState('SSPRESTA') ;
 lStEcrGen     := '' ;
 lStSqlCompl   := '' ;
 lStExo        := ctxExercice.QuelExoDate(StrToDate(GetControlText('DATEC'))) ;

 if lStatePresta = cbUnchecked then
  begin
   lStSqlCompl := ' AND  ( EC_CUTOFFDEB <>"' + UsDateTime(iDate1900) + '" ' + ' AND EC_CUTOFFFIN <>"' + UsDateTime(iDate1900) + '" ) ' ;
   if lState= cbChecked then
    lStEcrGen := ' AND EC_CUTOFFDATECALC<>"'+ USDATETIME(iDate1900) + '" '
     else
      if lState= cbUnChecked then
       lStEcrGen := ' AND EC_CUTOFFDATECALC="'+ USDATETIME(iDate1900) + '" ' ;

   lState    := GetCheckBoxState('EXHAUS') ;
   lStExhau  :=  ''  ;
   if lState= cbChecked then
    lStExhau := ' AND EC_CUTOFFFIN <"' + USDATETIME(StrToDate(GetControlText('DATEC'))) + '" '
     else
      if lState= cbUnChecked then
       lStExhau := ' AND EC_CUTOFFFIN >"' + USDATETIME(StrToDate(GetControlText('DATEC'))) + '" '  ;
  end
   else
    if lStatePresta = cbchecked then
     lStExhau  :=  ' AND(  ( ( EC_CUTOFFDEB IS NULL ) OR ( EC_CUTOFFDEB = "' + UsDateTime(iDate1900) + '" ) ) ) ' ;


// lStWhereGene   := ConvertitCaractereJokers(E_GENERAL, E_GENERAL_, 'E_GENERAL');

 AStSqlTobFListe := 'SELECT ' + CSqlTextFromList(FStListeChamps) +
                    ',E_DEBIT,E_CREDIT,G_SENS, G_NATUREGENE,G_LIBELLE,G_CONFIDENTIEL,E_MODESAISIE,G_GENERAL ' +
                    ', E_EXERCICE , E_JOURNAL, E_NUMLIGNE,E_NUMEROPIECE  ' +
                    ',E_NUMECHE , E_QUALIFPIECE ,E_QUALIFORIGINE, E_PERIODE , E_DEVISE,E_DATEMODIF,E_ETABLISSEMENT,EC_CUTOFFDATECALC, E_REFINTERNE,E_REFEXTERNE ' +
                    ' FROM GENERAUX , ECRITURE LEFT OUTER JOIN ECRCOMPL ON ( ' +
                    ' EC_JOURNAL=E_JOURNAL ' +
                    ' AND EC_EXERCICE=E_EXERCICE ' +
                    ' AND EC_DATECOMPTABLE=E_DATECOMPTABLE ' +
                    ' AND EC_EXERCICE=E_EXERCICE ' +
                    ' AND EC_NUMEROPIECE=E_NUMEROPIECE ' +
                    ' AND EC_NUMLIGNE=E_NUMLIGNE ' +
                    ' AND EC_QUALIFPIECE=E_QUALIFPIECE ) ' +
                    RecupWhereCritere(PageControl) +
                    lStExhau +
                    lStEcrGen +
                    ' AND ( E_EXERCICE="' + lStExo.Code + '" ) ' +
                    ' AND G_GENERAL=E_GENERAL ' +
                    ' AND E_DATECOMPTABLE<="' + USDATETIME(StrToDate(GetControlText('DATEC'))) + '" ' +
                    lStSqlCompl +
                    ' ORDER BY E_GENERAL,E_DATECOMPTABLE ' ;

 AStTriTobFListe  := 'E_GENERAL;E_DATECOMPTABLE' ;
 
end;


procedure TOF_CPMULCUTOFF.GetCellCanvasFListe(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState) ;
var
 lTOB : TOB ;
begin

 if (FListe.Row = ARow) or ( ACol <> 4 ) or ( ARow = 0 ) then Exit ;

 lTOB := GetO(FListe, ARow) ;
 if lTOB = nil then exit ;

 if lTOB.GetValue('G_SENS') = 'D' then
  Canvas.Font.Color := IIF(Pos('C', FListe.Cells[ACol, ARow]) > 0, ClRed, ClGreen)
   else
    if lTOB.GetValue('G_SENS') = 'C' then
     Canvas.Font.Color := IIF(Pos('C', FListe.Cells[ACol, ARow]) > 0, ClGreen, ClRed);

end;


procedure TOF_CPMULCUTOFF.OnDblClickFListe(Sender: TObject);
var
  AA     : TActionFiche ;
  M      : RMVT ;
  lQEcr  : TQuery ;
  lTOB   : TOB ;
begin

  lTOB := GetO(FListe, FListe.Row) ;

  if (not FListe.Focused) or ( lTOB = nil ) then Exit;

  lQEcr := nil;
  try
    lQEcr := OpenSql('SELECT * FROM ECRITURE WHERE ' +
      '(E_JOURNAL = "' + lTOB.GetValue('E_JOURNAL') + '") AND ' +
      '(E_EXERCICE = "' + lTOB.GetValue('E_EXERCICE') + '") AND ' +
      '(E_DATECOMPTABLE = "' + USDateTime(lTOB.GetValue('E_DATECOMPTABLE')) + '") AND ' +
      '(E_NUMEROPIECE = ' + IntToStr(lTOB.GetValue('E_NUMEROPIECE')) + ') AND ' +
      '(E_NUMLIGNE = ' + IntToStr(lTOB.GetValue('E_NUMLIGNE')) + ') AND ' +
      '(E_NumEche = ' + IntToStr(lTOB.GetValue('E_NUMECHE')) + ') AND ' +
      '(E_QualifPiece = "' + lTOB.GetValue('E_QUALIFPIECE') + '")', True);

    AA := taModif;
  //  if not FOkCreateModif then
  //    AA := taConsult;

    if (lQEcr.FindField('E_MODESAISIE').AsString <> '-') and
      (lQEcr.FindField('E_MODESAISIE').AsString <> '') then // Bordereau
      LanceSaisieFolio(lQEcr, AA)
    else
    begin
      // GC - 26/03/2003
      if TrouveSaisie(lQEcr, M, lQEcr.FindField('E_QUALIFPIECE').Asstring) then
      begin
        M.NumLigVisu := lQEcr.FindField('E_NUMLIGNE').AsInteger;
        LanceSaisie(lQEcr, AA, M);
      end;
    end;

  finally
    Ferme(lQEcr);
    BCherche.Click ;
  end;

end;


procedure TOF_CPMULCUTOFF.OnPopUpPopUpTraitement(Sender: TObject);
begin
 InitPopUp(True);
end;


procedure TOF_CPMULCUTOFF.InitPopUp( vActivation : Boolean) ;
var
 i    : integer ;
 lTOB : TOB ;
begin

 lTOB := GetO(FListe, FListe.Row);
 if not vActivation and ( lTOB = nil ) then exit ;

 if vActivation then
  for i := 0 to PopUpTraitement.Items.Count - 1 do
   begin

    if PopUpTraitement.Items[i].Name = 'MODIFECR' then
     begin
      PopUpTraitement.Items[i].OnClick := OnDblClickFListe;
      Continue;
     end;

    if PopUpTraitement.Items[i].Name = 'ECRCPTCPT' then
     begin
      PopUpTraitement.Items[i].Visible := false ;
      Continue;
     end;

     // Separator
     if PopUpTraitement.Items[i].Name = 'SAISIEBOR' then
      begin
       PopUpTraitement.Items[i].OnClick := SaisieBorClick;
       Continue;
      end;
   {$IFNDEF EAGLCLIENT}
    if PopUpTraitement.Items[i].Name = 'SAISIEPIECE' then
     begin
      PopUpTraitement.Items[i].OnClick := SaisieEcrClick;
      Continue;
     end;
     {$ENDIF}

    if PopUpTraitement.Items[i].Name = 'JUSTIFCOMPTE' then
     begin
      PopUpTraitement.Items[i].OnClick := OnClickEtatJustifCompte ;
      Continue;
     end;

    if PopUpTraitement.Items[i].Name = 'JUSTIFPIECE' then
     begin
      PopUpTraitement.Items[i].OnClick := OnClickBImprimer ;
      Continue;
     end;

   end;

 // PopUp Utilitaire
 for i := 0 to PopUpUtilitaire.Items.Count - 1 do
  begin

    if ( PopUpUtilitaire.Items[i].Name = 'INFOCOMP' ) then
     begin
      PopUpUtilitaire.Items[i].OnClick := OnClickInfosComp ;
      continue ;
     end ;

    if PopUpUtilitaire.Items[i].Name = 'PARAMGENERAL' then
     begin
      if vActivation then
       PopUpUtilitaire.Items[i].OnClick := OnClickCompteGeneral
        else
         PopUpUtilitaire.Items[i].Enabled :=  not EstConfidentiel(lTOB.GetValue('G_CONFIDENTIEL')) ;
      Continue  ;
     end;

  end;

   // PopUp Utilitaire
 for i := 0 to PopUpEtat.Items.Count - 1 do
  begin

    if ( PopUpUtilitaire.Items[i].Name = 'CUT1' ) then
     begin
      PopUpUtilitaire.Items[i].OnClick := OnClickEtatJustifCompte ;
      continue ;
     end ;

    if PopUpUtilitaire.Items[i].Name = 'CUT2' then
     begin
      if vActivation then
       PopUpUtilitaire.Items[i].OnClick := OnClickCompteGeneral ;
      Continue  ;
     end;

  end;


end ;


procedure TOF_CPMULCUTOFF.SaisieBorClick(Sender: TObject);
begin
 SaisieFolio(taModif);
 BCherche.Click;
end;

{$IFNDEF EAGLCLIENT}
procedure TOF_CPMULCUTOFF.SaisieEcrClick(Sender: TObject);
var OldSC: Boolean;
begin
  OldSC := VH^.BouclerSaisieCreat;
  VH^.BouclerSaisieCreat := False;
  MultiCritereMvt(taCreat, 'N', False);
  VH^.BouclerSaisieCreat := OldSC;
  BCherche.Click;
end;
{$ENDIF}

procedure TOF_CPMULCUTOFF.OnPopUpPopF11(Sender: TObject);
begin
 InitPopUp(True);
 inherited; // Ajoute les élements visibles des menus de Ecran dans PopF11
end;

procedure TOF_CPMULCUTOFF.OnPopUpPopUpUtilitaire(Sender: TObject);
begin
  InitPopUp(True);
end;

procedure TOF_CPMULCUTOFF.OnPopUpEtat(Sender: TObject);
begin
  InitPopUp(True);
end;

{***********A.G.L.***********************************************
Auteur  ...... : LG
Créé le ...... : 14/02/2007
Modifié le ... : 20/04/2007
Description .. : - FB 18700 - 14/02/2007 - on ne gere plus le bouton valider 
Suite ........ : ici, mais ds le formkeydown
Suite ........ : - FB 19770 - 20/04/2007 - on supprime cette gestion
Mots clefs ... : 
*****************************************************************}
{procedure TOF_CPMULCUTOFF.OnClickSansPresta ( Sender : TObject ) ;
var
 lBoActif : boolean ;
begin  exit ;
 lBoActif := GetCheckBoxState('SSPRESTA') = cbUnChecked ;
 SetControlEnabled('DATEC'         , lBoActif ) ;
 SetControlEnabled('JOURNAL'       , lBoActif ) ;
 SetControlEnabled('QUALIFPIECE'   , lBoActif ) ;
 SetControlEnabled('EXHAUS'        , lBoActif ) ;
end ;   }


{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 17/04/2007
Modifié le ... :   /  /    
Description .. : - 17/04/2007 - LG - FB 19843 - on renseigne la cle des info 
Suite ........ : compl avant de l'enregistrer en base, sion on cree un ligne 
Suite ........ : ss exercice, journal
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPMULCUTOFF.OnClickInfosComp(Sender: TObject);
var
  OBM    : TOBM;
  RC     : R_COMP;
  AA     : TActionFiche;
  ModBN  : Boolean;
  lQEcr  : TQuery;
  lTOB   : TOB ;
begin

 lTOB := _RecupEcrselect(Fliste) ;

 if lTOB = nil then Exit;

  lQEcr := nil;
  OBM   := nil;
  try
    lQEcr := OpenSql('SELECT * FROM ECRITURE WHERE ' +
             '(E_JOURNAL = "' + lTOB.GetValue('E_JOURNAL') + '") AND ' +
             '(E_EXERCICE = "' + lTOB.GetValue('E_EXERCICE') + '") AND ' +
             '(E_DATECOMPTABLE = "' + USDateTime(lTOB.GetValue('E_DATECOMPTABLE')) + '") AND ' +
             '(E_NUMEROPIECE = ' + IntToStr(lTOB.GetValue('E_NUMEROPIECE')) + ') AND ' +
             '(E_NUMLIGNE = ' + IntToStr(lTOB.GetValue('E_NUMLIGNE')) + ') AND ' +
             '(E_NUMECHE = ' + IntToStr(lTOB.GetValue('E_NUMECHE')) + ') AND ' +
             '(E_QUALIFPIECE = "' + lTOB.GetValue('E_QUALIFPIECE') + '")', False);

    AA := taModif;

    OBM := TOBM.Create(EcrGen, '', False);
    OBM.ChargeMvt( lQEcr );
    Ferme( lQEcr );

    if OBM <> nil then
    begin
      FInfo.LoadCompte( lTOB.GetValue('E_GENERAL') ) ;
      RC.StLibre     := '---CUTXXXXXXXXXXXXXXXXXXXXXXXX' ;
      RC.StComporte  := '--XXXXXXXX' ;
      RC.CutOffPer   := FInfo.GetString('G_CUTOFFPERIODE') ;
      RC.CutOffEchue := FInfo.GetString('G_CUTOFFECHUE') ;
      ModBN          := True;
      RC.Conso       := True;
      RC.Attributs   := False;
      RC.MemoComp    := nil;
      RC.Origine     := -1;
      RC.TOBCompl    := CSelectDBTOBCompl(OBM,nil) ;
      if SaisieComplement(TOBM(lTOB), EcrGen, AA, ModBN, RC, False, True) then
      begin
        CMAJTOBCompl(OBM) ; // FB 19843
        lTOB.SetAllModifie(true) ;
        lTOB.UpdateDb;
        RC.TOBCompl.InsertOrUpdateDB(false) ;
        BCherche.Click;
      end;
    end;

  finally
    Ferme(lQEcr);
    RC.TOBCompl.Free ;
    FreeAndNil(OBM);
    FreeAndNil(lTOB);
  end;

end;


procedure TOF_CPMULCUTOFF.OnClickCompteGeneral(Sender: TObject);
var
 lAction : TActionFiche;
 lTobEcr : TOB ;
begin

 lTobEcr := GetO(FListe, FListe.Row);
 if lTobEcr = nil then exit ;

 lAction := TaModif;

 if not ExJaiLeDroitConcept(TConcept(ccGenModif), False) then
  lAction := taConsult;

 FicheGene(nil, '', lTobEcr.getValue('E_GENERAL'), lAction, 0);
 BCherche.Click;

end;

procedure TOF_CPMULCUTOFF.OnElipsisClickE_General(Sender: TObject);
begin
  LookUpList( THEdit(Sender),
              'Comptes généraux',
              'GENERAUX',
              'G_GENERAL' ,
              'G_LIBELLE',
              CGenereSQLConfidentiel('G'),
              'G_GENERAL' ,
              True,
              1 );
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 21/09/2005
Modifié le ... : 14/02/2007
Description .. : - LG - 21/09/2005 - FB 16679 - on ne complete aps en cas 
Suite ........ : de presence d'un caractere joker
Suite ........ : - LG - 14/02/2007 - FB 18574 - ajout du traitemetn du * en 
Suite ........ : caractere joker
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPMULCUTOFF.OnExitE_General(Sender : TObject) ;
var
 lStVal : string ;
begin
 lStVal := GetControlText('G_GENERAL') ;
 if ( Pos('%',lStVal ) > 0 ) or ( Pos('*',lStVal ) > 0  ) then exit ;
 if trim(lStVal) <> '' then
  THEdit(GetControl('G_GENERAL', True)).Text := BourreEtLess(lStVal, fbGene) ;
 lStVal := GetControlText('G_GENERAL_') ;
 if ( Pos('%',lStVal ) > 0 ) or ( Pos('*',lStVal ) > 0  ) then exit ;
 if trim(lStVal) <> '' then
  THEdit(GetControl('G_GENERAL_', True)).Text := BourreEtLess(lStVal, fbGene) ;   
end;

procedure TOF_CPMULCUTOFF.OnExitDateC(Sender : TObject);
begin
 if not CControleDateBor(StrToDate(GetControlText('DATEC')),FInfo.Exercice,true) then
  begin
   if VH^.CPExoRef.Code<>'' then
    SetControlText('DATEC',DateToStr(VH^.CPExoRef.Fin) )
     else
      SetControlText('DATEC',DateToStr(VH^.EnCours.Fin) ) ;
  end ;
end;



{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 27/05/2005
Modifié le ... : 02/03/2007
Description .. : - 27/05/2005 - LG - FB 15937 - Ne pas générer de 
Suite ........ : mouvements lorsque les zones montant CCA ou montant 
Suite ........ : PCA sont à 0
Suite ........ : - 02/03/2007 - LG - bug qd on selectionnait l'essemble des 
Suite ........ : lignes
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPMULCUTOFF.OnClickBGENE( Sender : TObject) ;
var
 i          : integer ;
 lTOBSelect : TOB ;
 lTOB       : TOB ;
 lBoResult  : boolean ;
begin

 if (  FListe.nbSelected = 0 ) then
  begin
   PGIInfo('Aucune ligne de sélectionnée') ;
   Ecran.ModalResult := mrNone; 
   exit ;
  end ;


 if ( PGIAsk('Confirmez-vous la génération des écritures ?',Ecran.Caption) = mrNo ) then
  begin                       
   Ecran.ModalResult := mrNone ;
   exit ;
  end ;

 Ecran.ModalResult := MrYes ;

 lTOBSelect := TOB.Create('',nil,-1) ;

 for i := 1 to FListe.RowCount - 1 do
  if FListe.IsSelected(i) then
   begin
     lTOB := GetO(FListe,i) ;
     if lTOB = nil then exit ;
     if Arrondi(lTOB.GetValue('MCCA')+lTOB.GetValue('MPCA'), 2 ) = 0 then continue ;
     lTOB.ChangeParent(lTOBSelect,-1) ;
   end ;

 InitMoveProgressForm(Ecran,'Génération en cours...','Génération en cours',lTOBSelect.Detail.Count,true,false) ;

 try

  lBoResult := FTCutOff.Execute(lTOBSelect,StrToDate(GetControlText('DATEC')),GetControlText('JOURNAL') , GetControlText('QUALIFPIECE') ) ;
  if not lBoResult and ( V_PGI.LastSQLError <> '' ) then
    PGIError(V_PGI.LastSQLError) ;

  if lBoResult and ( FTCutOff.CR.count > 0 ) then
   begin
    AfficheCRIntegration (FTCutOff.CR); // fct de libI\ImRapInt.pas
    FTCutOff.RAZFLC ;
   end; // if


 finally
  FreeAndNil(lTOBSelect) ;
  FiniMoveProgressForm ;
  if FListe.Row <> FListe.RowCount - 1 then
   FListe.ClearSelected ;
  BCherche.Click ;
 end ;

end;

procedure TOF_CPMULCUTOFF.OnErrorTOB( sender : TObject ; Error : TRecError);
begin
 if ( trim(Error.RC_Message) <> '' ) then
  PGIInfo(Error.RC_Message + #10#13 + Error.RC_Methode , Ecran.Caption )
   else
    if Error.RC_Error <> RC_PASERREUR  then
      FMessCompta.Execute(Error.RC_Error) ;
end;


procedure TOF_CPMULCUTOFF.FListePostDrawCell( ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState );
begin

 if csDestroying in Ecran.ComponentState then Exit ;

 if ( ARow < FListe.FixedRows ) or ( ARow > ( FListe.RowCount - 1 ) ) then Exit ;

 try

 // gestion des cases debit et crédit - une seul des deux doit être renseignée
 if (((ACol = 9) or (ACol = 10)) and
     (Trim(FListe.Cells[ACol, ARow]) = '')) or
     (FListe.Cells[ACol, ARow] = '0') then
  begin
   FListe.PostDrawCell  := nil ; // on debranche l'évènement lors du dessin de la grille
   SetGridGrise(ACol, ARow, FListe) ;
   FListe.PostDrawCell  := FListePostDrawCell ;
  end;

 except
  on E : exception do
   begin
     FListe.PostDrawCell := nil;
     PGIBox( 'Problème lors de l''affichage des données !' + #13#10 +
             E.Message ,
             Ecran.Caption);
     V_PGI.IoError := oeSaisie;
   end; // on
 end;

end ;


{***********A.G.L.***********************************************
Auteur  ...... : LG
Créé le ...... : 24/01/2007
Modifié le ... :   /  /    
Description .. : - FB 16914 - 24/01/2007 - Ajout des journaux estra
Suite ........ : comtable
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPMULCUTOFF.QUALIFPIECEChange ( Sender : TObject );
begin
 FListe.ClearSelected ;
 FListe.Refresh ;
 if GetControlText('QUALIFPIECE') = 'S' then
  E_JOURNAL.Plus := ' and ( J_NATUREJAL="REG" or J_NATUREJAL="OD" OR J_NATUREJAL="EXT" ) '
   else
    E_JOURNAL.Plus := ' and ( J_NATUREJAL="REG" or J_NATUREJAL="OD" OR J_NATUREJAL="EXT" ) ' +
                      ' or ( ( J_MODESAISIE="BOR" or J_MODESAISIE="LIB" ) and  (J_NATUREJAL="REG" or J_NATUREJAL="OD" OR J_NATUREJAL="EXT" ) ) ';

end; 

{***********A.G.L.***********************************************
Auteur  ...... : L Gendreau
Créé le ...... : 14/09/2005
Modifié le ... :   /  /
Description .. : - 14/09/2005 - FB 16680 - La génération d'une écriture de
Suite ........ : situation (ou simulation) avec un journal "Bordereau" !
Suite ........ : corrigé
Mots clefs ... :
*****************************************************************}
procedure TOF_CPMULCUTOFF.JOURNALChange ( Sender : TObject );
var
 lStJal : string ;
begin
 lStJal := E_JOURNAL.Value ;
 if FInfo.LoadJournal(lStJal) then
  begin
   if ( FInfo.Journal.GetValue('J_MODESAISIE') = 'BOR' ) or ( FInfo.Journal.GetValue('J_MODESAISIE') = 'LIB' ) then
    begin
     THValComboBox(GetControl('QUALIFPIECE', True)).DataType := 'TTQUALPIECE' ;
     THValComboBox(GetControl('QUALIFPIECE', True)).Plus     := ' AND CO_CODE="N" ' ;
    end
     else
      begin
       THValComboBox(GetControl('QUALIFPIECE', True)).DataType := 'TTQUALPIECE' ;
       THValComboBox(GetControl('QUALIFPIECE', True)).Plus     := '' ;
      end ;
  end ;
end;

function TOF_CPMULCUTOFF.BeforeLoad: boolean;
var
 lStDate : string ;
begin
 lStDate := GetControlText('DATEC') ;
 if not isValidDate(lStDate) or ( not CControleDateBor(StrToDate(lStDate),FInfo.Exercice,true) ) then
  begin
   if VH^.CPExoRef.Code<>'' then
    SetControlText('DATEC',DateToStr(VH^.CPExoRef.Fin) )
     else
      SetControlText('DATEC',DateToStr(VH^.EnCours.Fin) ) ;
  end ;
    result := true ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 28/06/2007
Modifié le ... :   /  /    
Description .. : - LG - 28/06/2007 - FB 20828 - on enlevel les colformats 
Suite ........ : pour l'objet d'edtion auto de la grille
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPMULCUTOFF.OnClickBImprimer (Sender : Tobject) ;
var
 l1 : string ;
 l2 : string ;
 l3 : string ;
begin
 l1                   := FListe.ColFormats[3] ;
 l2                   := FListe.ColFormats[4] ;
 l3                   := FListe.ColFormats[11] ;
 FListe.ColFormats[3] := '' ;
 FListe.ColFormats[4] := '' ;
 FListe.ColFormats[11]:= '' ;
 inherited ;
 FListe.ColFormats[3] := l1 ;
 FListe.ColFormats[4] := l2 ;
 FListe.ColFormats[11]:= l3 ;
 RefreshPclPge ;
end ;


{ TCutOff }

constructor TCutOff.Create( vInfoEcr : TInfoEcriture );
begin
 inherited ;
 FTOBN1 := TOB.Create('',nil,-1) ;
 FCR    := TList.Create ;
end;

destructor TCutOff.Destroy;
begin
 FreeAndNil(FTOBN1) ;
 RAZFLC ;
 FreeAndNil(FCR) ;
 inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gendreau Laurent
Créé le ...... : 04/07/2007
Modifié le ... :   /  /    
Description .. : - LG - 04/07/2007 - FB 20813 - la validation est sorti ds la 
Suite ........ : transaction
Mots clefs ... : 
*****************************************************************}
function TCutOff.Execute( vTOB : TOB ; vDateC : TDateTime ; vJournal,vQualifPiece : string  ) : boolean ;
var
 i : integer ;
begin

 FTOBN3       := vTOB ;
 FDateC       := vDateC ;
 FJournal     := vJournal ;
 FQualifPiece := vQualifPiece ;

 FTOBN1.ClearDetail ;

 AjouteN3aN1 ;

 for i := 0 to FTOBN1.Detail.Count - 1 do
  begin
   FTOB := FTOBN1.Detail[i] ;
   FTOBGene := TOB.Create('',nil,-1) ;
   if Construire then
    Transactions(Generer,0) ;
   FTOBGene.Free ;
  end ; // if

 result := V_PGI.IOError = oeOk ;

end ;

function TCutOff.ChargeInfo : boolean ;
begin

 result := false ;

 if ( FTOB.Detail =  nil ) or ( FTOB.Detail.Count = 0 ) then
   begin
    NotifyError('Problème lors de la récupérationd des informations de la grille','','TCutOff.ChargeInfo') ;
    exit ;
   end ;

  if Info.Devise.Load( [FTOB.Detail[0].GetValue('E_DEVISE')] ) = - 1 then
   begin
    NotifyError(RC_DEVISE,'','TCutOff.ChargeInfo') ;
    exit ;
   end ;

  if not Info.LoadJournal(GetJournal(FTOB.Detail[0])) then
   begin
    NotifyError(RC_JALINEXISTANT,'','TCutOff.ChargInfo') ;
    exit ;
   end ;

  result := true ;

end ;

function TCutOff.MAJEcriture : boolean ;
var
 lInNB : integer ;
 i     : integer ;
 lTOB  : TOB ;
begin

  result := false ;

 // on mets a jour les base de depart
  for i := 0 to FTOB.Detail.Count - 1 do // N2
   begin

    BeginTrans ;

    try

    lTOB  := FTOB.Detail[i] ;
    lInNb := ExecuteSQL('UPDATE ECRITURE SET E_DATEMODIF="'+UsTime(NowH) + '" ' +
                        ' WHERE ' +
                        'E_JOURNAL = "'        + lTOB.GetValue('E_JOURNAL')                      + '" AND ' +
                        'E_EXERCICE = "'       + lTOB.GetValue('E_EXERCICE')                     + '" AND ' +
                        'E_DATECOMPTABLE = "'  + USDateTime(lTOB.GetValue('E_DATECOMPTABLE'))    + '" AND ' +
                        'E_NUMEROPIECE = '     + IntToStr(lTOB.GetValue('E_NUMEROPIECE'))        + '  AND ' +
                        'E_NUMLIGNE = '        + IntToStr(lTOB.GetValue('E_NUMLIGNE'))           + '  AND ' +
                        'E_NUMECHE = '         + IntToStr(lTOB.GetValue('E_NUMECHE'))            + '  AND ' +
                        'E_QUALIFPIECE = "'    + lTOB.GetValue('E_QUALIFPIECE')                  + '" AND ' +
                        'E_DATEMODIF="'        + UsTime(lTOB.GetValue('E_DATEMODIF'))+'" ' ) ;
    if lInNb = 1 then
    lInNb := ExecuteSQL('UPDATE ECRCOMPL SET EC_CUTOFFDATECALC="'+ UsDateTime(FDateC) + '" ' +
                        ' WHERE ' +
                        'EC_JOURNAL = "'        + lTOB.GetValue('E_JOURNAL')                      + '" AND ' +
                        'EC_EXERCICE = "'       + lTOB.GetValue('E_EXERCICE')                     + '" AND ' +
                        'EC_DATECOMPTABLE = "'  + USDateTime(lTOB.GetValue('E_DATECOMPTABLE'))    + '" AND ' +
                        'EC_NUMEROPIECE = '     + IntToStr(lTOB.GetValue('E_NUMEROPIECE'))        + '  AND ' +
                        'EC_NUMLIGNE = '        + IntToStr(lTOB.GetValue('E_NUMLIGNE'))           + '  AND ' +
                        'EC_NUMECHE = '         + IntToStr(lTOB.GetValue('E_NUMECHE'))            + '  AND ' +
                        'EC_QUALIFPIECE = "'    + lTOB.GetValue('E_QUALIFPIECE')                  + '" ' ) ;

    finally
     CommitTrans ;
    end ;

    if lInNB <> 1 then exit ;

   end ; // for

  result := true ;

end ;

function _AffecteLib ( vBoAuDebit : boolean ) : string ;
begin
 if vBoAuDebit then
  result := 'CCA du '
   else
    result := 'PCA du ' ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 22/07/2005
Modifié le ... :   /  /    
Description .. : - LG - 22/07/2005 - FB 16041 - reaffectation de la nature
Suite ........ : pour les lignes d'analytiq
Mots clefs ... : 
*****************************************************************}
procedure TCutOff.TransfertAna ( vTOBEcr,vTOBCut : TOB ) ;
var
 i    : integer ;
 j    : integer ;
begin

 CChargeAna(vTOBEcr) ;
 AlloueAxe(vTOBCut) ;

 for i := 0 to vTOBEcr.Detail.Count - 1 do
  begin
   for j := 0 to vTOBEcr.Detail[i].Detail.Count - 1 do
    begin
     vTOBEcr.Detail[i].Detail[j].PutValue('Y_NATUREPIECE','OD') ;
     vTOBEcr.Detail[i].Detail[j].PutValue('Y_JOURNAL',GetJournal(FTOB.Detail[0])) ;
    end ;
   vTOBCut.Detail[i].Dupliquer(vTOBEcr.Detail[i],true,true) ;
  end ;

 RecalculProrataAnalNEW('Y',vTOBCut,0,Info.Devise.Dev) ;

end ;


{***********A.G.L.***********************************************
Auteur  ...... : LG
Créé le ...... : 07/01/2008
Modifié le ... :   /  /    
Description .. : - LG - 07/01/2008 - FB 21896 - correction de la generation 
Suite ........ : de l'analytiq
Mots clefs ... : 
*****************************************************************}
function TCutOff.Construire : boolean ;
var
 i               : integer ;
 lRdMontant      : double ;
 lTOB            : TOB ;
 lTOBEcr         : TOB ;
 lTOBCut         : TOB ;
 lStLib          : string ;
// lTOBGene        : TOB ;
// lStCompte       : string ;
 lRecError       : TRecError ;
 lStModeSaisie   : string ;
 lStNatureGene   : string ;
 lStContrepartie : string ;
 lStEta          : string ;
begin

 result := false ;

 if not ChargeInfo then exit ;

 lStModeSaisie := Info.Journal.GetValue('J_MODESAISIE') ;

 // GCO - 30/05/2006 - FQ 17837 - Type d'écritures à générer ( N ou S ou U )
 lTOBCut       := TOB.Create('',nil,-1) ;
 lStEta        := FTOB.Detail[0].GetValue('E_ETABLISSEMENT') ;

 try

 for i := 0 to FTOB.Detail.Count - 1 do // N2
  begin

   //1er ligne
   lTOBEcr := _RecupEcrSelect(nil,FTOB.Detail[i]) ;
   lTOB := TOBM.Create( EcrGen,'',true,FTOBGene);
   lTOB.Dupliquer(lTOBEcr, false,true) ;
   lTOB.PutValue('E_NUMLIGNE'         , i + 1 ) ;
   lTOB.PutValue('E_ETABLISSEMENT'    , lStEta ) ;
   lTOB.PutValue('E_NUMEROPIECE'      , -1 ) ;
   lTOB.PutValue('E_NUMGROUPEECR'     , 0 ) ;
   lTOB.PutValue('E_NATUREPIECE'      , 'OD' ) ;
   lTOB.PutValue('E_GENERAL'          , FTOB.Detail[i].GetValue('E_GENERAL') ) ;
   lTOB.PutValue('E_JOURNAL'          , GetJournal(FTOB.Detail[i]) ) ;
   lTOB.PutValue('E_MODESAISIE'       , lStModeSaisie ) ;
   lTOB.PutValue('E_EXERCICE'         , QuelExoDT(FDateC) ) ;
   lTOB.PutValue('E_QUALIFORIGINE'    , 'CUT' ) ;
   lTOB.PutValue('E_QUALIFPIECE'      , GetQualifPiece ) ;
   CRemplirDateComptable( lTOB , FDateC ) ;
   lStLib := _AffecteLib( FTOB.Detail[i].GetValue('MCCA') > 0 ) ;
   lStLib := lStLib + DateToStr(FDateC+1) + ' au ' + DateToStr(FTOB.Detail[i].GetValue('EC_CUTOFFFIN')) ;
   lTOB.PutValue('E_LIBELLE'          , lStLib ) ;
   lTOB.PutValue('E_IO'               , 'X') ;
   lTOB.PutValue('E_REFPOINTAGE'      , '') ;
   lTOB.PutValue('E_DATEPOINTAGE'     , iDate1900) ;

   Info.LoadCompte ( FTOB.Detail[i].GetValue('E_GENERAL') ) ;
   CSetMontants(lTOB,FTOB.Detail[i].GetValue('MPCA'),FTOB.Detail[i].GetValue('MCCA'),Info.Devise.Dev,true) ;
   lStNatureGene := Info.GetString('G_NATUREGENE') ;

   if lStNatureGene = 'CHA' then
    lRdMontant := FTOB.Detail[i].GetValue('MCCA')
     else
      lRdMontant := FTOB.Detail[i].GetValue('MPCA') ;

     // on recupere le compte de cutoff qui est aussi le compte de contrepartie
   lStContrepartie := trim(Info.GetString('G_CUTOFFCOMPTE')) ;

   if trim(lStContrepartie) = '' then
    begin
    // on prends la parametre par defaut
     if lStNatureGene = 'CHA' then
      lStContrepartie := GetParamSocSecur('SO_CPCCA','')
       else
        lStContrepartie := GetParamSocSecur('SO_CPPCA','') ;
    end ;

   lTOB.PutValue('E_CONTREPARTIEGEN'  , lStContrepartie ) ;

   if Info.Compte.IsVentilable then
    begin
     FTOB.Detail[i].AddChampSupValeur('E_ANA','X',false) ; // FB 21896
     TransfertAna(FTOB.Detail[i],lTOB) ;
     lTOB.PutValue('E_ANA','X' ) ;
    end ;// if

   if Info.Compte.IsLettrable then
    CRemplirInfoLettrage(lTOB)
     else
      CSupprimerInfoLettrage(lTOB);

   if Info.Compte.IsPointable then
    CRemplirInfoPointage(lTOB)
     else
      begin
       lTOB.PutValue('E_REFPOINTAGE'  , '') ;
       lTOB.PutValue('E_DATEPOINTAGE' , iDate1900) ;
      end ;

   CGetTVA(lTOB,Info ) ;
   CGetEch(lTOB,Info) ;

   // 2eme ligne

   Info.LoadCompte( lStContrepartie ) ;
   // on recherche s'il on a deja generer une ligne sur ce compte
   lTOB := lTOBCut.FindFirst(['E_GENERAL'],[lStContrepartie],true) ;
   if lTOB = nil then
    begin
     lTOB := TOBM.Create( EcrGen,'',true,lTOBCut) ;
     lTOB.PutValue('E_GENERAL' , lStContrepartie) ;
    end
     else // si oui on additionne les montants
      lRdMontant := Arrondi(lRdMontant + lTOB.GetValue('E_DEBIT') - lTOB.GetValue('E_CREDIT') , 0 ) ;

   if lStNatureGene = 'CHA' then
    CSetMontants(lTOB,lRdMontant,0,Info.Devise.Dev,true)
     else
      CSetMontants(lTOB,0,lRdMontant,Info.Devise.Dev,true) ;

   lTOB.PutValue('E_NUMLIGNE'         , -1 ) ;
   lTOB.PutValue('E_NUMEROPIECE'      , -1 ) ;
   lTOB.PutValue('E_NUMGROUPEECR'     , 0 ) ;
   lTOB.PutValue('E_NATUREPIECE'      , 'OD' ) ;
   lTOB.PutValue('E_ETABLISSEMENT'    , lStEta ) ;
   lTOB.PutValue('E_GENERAL'          , lStContrepartie  ) ;
   lTOB.PutValue('E_EXERCICE'         , QuelExoDT(FDateC) );
   lTOB.PutValue('E_JOURNAL'          , GetJournal(FTOB.Detail[0]) ) ;
   lTOB.PutValue('E_QUALIFPIECE'      , GetQualifPiece ) ;
   lTOB.PutValue('E_MODESAISIE'       , lStModeSaisie ) ;
   CRemplirDateComptable( lTOB        , FDateC ) ;
   lTOB.PutValue('E_AUXILIAIRE'       , '' ) ;
   lTOB.PutValue('E_LIBELLE'          , Info.GetString('G_LIBELLE') ) ;
   lTOB.PutValue('E_IO'               , 'X') ;
   lTOB.PutValue('E_QUALIFORIGINE'    , 'CUT' ) ;
   lTOB.PutValue('E_REFPOINTAGE'      , '') ;
   lTOB.PutValue('E_DATEPOINTAGE'     , iDate1900) ;
   lTOB.PutValue('E_CONTREPARTIEGEN'  , FTOB.Detail[i].GetValue('E_GENERAL') ) ;

   if Info.Compte.IsLettrable then
    CRemplirInfoLettrage(lTOB)
     else
      CSupprimerInfoLettrage(lTOB);

   if Info.Compte.IsPointable then
    CRemplirInfoPointage(lTOB) ;

   CGetTVA(lTOB,Info ) ;
   CGetEch(lTOB,Info) ;

   lTOBEcr.Free ;

  end ; // for

 while lTOBCut.Detail.Count <> 0 do
  lTOBCut.Detail[0].ChangeParent(FTOBGene,-1) ;

 if ( FTOBGene.Detail = nil ) or ( FTOBGene.Detail.Count = 0 ) then exit ;

 if CEquilibrePiece(FTOBGene) then
  begin
//   CAffectCompteContrePartie(lTOBGene,Info) ;
   CAffectRegimeTva(FTOBGene) ;
  end;

 lRecError := CIsValidSaisiePiece ( FTOBGene , Info , true) ;
 if ( lRecError.RC_Error <> RC_PASERREUR ) then
  begin
   OnError(nil,lRecError) ;
   exit ;
  end ;

 result := true ;

 finally
  FreeAndNil(lTOBCut) ;
 end; // try

end ;



{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 23/05/2005
Modifié le ... : 19/02/2007
Description .. : - FB 15889 - LG - 23/05/2005 - L'écriture étant générale, il
Suite ........ : faut cumuler sur les mêmes numéros de compte
Suite ........ : - FB 15904 - LG - 23/05/2005 - Nature de la pièce =
Suite ........ : OPERATIONS DIVERSES
Suite ........ : - FB 16042 - LG - 16/06/2005 - ajout du lettrage sur les
Suite ........ : compte de cutoff
Suite ........ : - FB 15904 - LG - 25/08/2005 - le champs e_qualifpiece
Suite ........ : n'etait affecte que sur la éeme ligne
Suite ........ : - FB 16818 - LG - 05/10/2005 - suppression du controle sur
Suite ........ : la qualifpiece , fait ds le noyau
Suite ........ : - FB 15890 - LG - 12/09/2005 - correction du calcul pour
Suite ........ : les comptes de charges
Suite ........ : - FB 19382 - 03/01/2007 - affectation du
Suite ........ : champs e_modepaie  et suppression de la ref de pointage
Suite ........ : - FB 19693 - 19/02/2007 - on generait des pieces sur
Suite ........ : plusieurs etablissement
Mots clefs ... :
*****************************************************************}
procedure TCutOff.Generer ;
begin

 V_PGI.IOError := oeSaisie ;

 if CEnregistreSaisie(FTOBGene, True, False, True, Info, (FQualifPiece = 'N') ) then
  begin
   MoveCurProgressForm('Intégration en cours...') ;
   if MAJEcriture then
    StockeLesEcrituresIntegrees(FTOBGene)
     else
      exit ;
  end ;

 V_PGI.IOError := oeOk ;

end ;


function TCutOff.MakeCrit( vTOB : TOB ) : string ;
begin
 result := vTOB.GetValue('E_EXERCICE')                 +
           vTOB.GetValue('E_ETABLISSEMENT')            +
           GetJournal(vTOB)                            +
           vTOB.GetValue('E_DEVISE')                   ;

end ;

procedure TCutOff.AjouteN3aN1 ;
var
 lStCrit       : string ;
 lInIndex      : integer ;
 lTOBN2        : TOB ;
begin

 lInIndex := 0 ;

 while  FTOBN3.Detail.Count <> 0 do
  begin

    lStCrit := MakeCrit( FTOBN3.Detail[lInIndex] ) ;
    lTOBN2  := FTOBN1.FindFirst(['CRIT'],[lStCrit],false) ;
    if lTOBN2 = nil then
     begin
       lTOBN2 := TOB.Create( '', FTOBN1 ,-1) ;
       lTOBN2.AddChampSupValeur('CRIT',lStCrit) ;
       lTOBN2.AddChampSupValeur('NIV',2) ;
     end ; // if
    FTOBN3.Detail[lInIndex].ChangeParent(lTOBN2,-1) ;

  end; // while

end;

procedure TCutOff.RAZFLC ;
var
 i       : integer ;
 ARecord : TCRPiece ;
begin
 for i := 0 to (FCr.Count - 1) do
  begin
   ARecord := FCR.Items[i];
   ARecord.Free;
  end; // for
 FCR.Clear ;
end;

procedure TCutOff.StockeLesEcrituresIntegrees( vTOB : TOB) ;
var
 lLigneCR : TCRPiece ; // defini ds ImRapInt
 lTOB     : TOB ;
begin
 if vTOB.Detail.Count > 0 then
  begin
   lLigneCR             := TCRPiece.Create ;
   lTOB                 := vTOB.Detail[0] ;
   lLigneCR.NumPiece    := lTOB.GetValue('E_NUMEROPIECE') ;
   lLigneCR.Journal     := lTOB.GetValue('E_JOURNAL') ;
   lLigneCR.QualifPiece := lTOB.GetValue('E_QUALIFPIECE') ;
   lLigneCR.Date        := lTOB.GetValue('E_DATECOMPTABLE') ;
   lLigneCR.Debit       := vTOB.Somme('E_DEBIT', ['E_NUMEROPIECE'],[lLigneCR.NumPiece],TRUE);
   lLigneCR.Credit      := vTOB.Somme('E_CREDIT',['E_NUMEROPIECE'],[lLigneCR.NumPiece],TRUE);
   FCR.Add (lLigneCR);
  end; // for
end;

function TCutOff.GetJournal ( vTOB : TOB ) : string ;
begin
 result := FJournal ;
 if result = '' then
  result := vTOB.GetValue('E_JOURNAL') ;
end ;

function TCutOff.GetQualifPiece : string ;
begin
 result := FQualifPiece ;
 if result = '' then
  result := 'N' ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 29/05/2007
Modifié le ... : 26/06/2007
Description .. : - FB 20403 - LG - corection du parametrage du 
Suite ........ : LanceEtatTOB
Suite ........ : - LG - 26/06/2007 - FB 20815 - tri par date comptable
Mots clefs ... :
*****************************************************************}
procedure TOF_CPMULCUTOFF.OnClickEtatJustifCompte(Sender: TObject) ;
var
 lBoOldCascadeForms : boolean ;
 lStRetour          : string ;
 lTOB               : TOB ;
 lTOBDup            : TOB ;
 i                  : integer ;
 lStRup             : string ;
 lStSaut            : string ;
 lStCrit            : string ;
begin

 if ATobFListe.Detail.Count = 0 then exit ;

 lBoOldCascadeForms := V_PGI.CascadeForms ;
 V_PGI.CascadeForms := false ;
 try
  lStRetour := AGLLanceFiche('CP', 'CPCUTOFFPARAM', '', '', '' ) ;
 finally
  V_PGI.CascadeForms := lBoOldCascadeForms ;
 end ;

 if lStRetour = '' then exit ;

 lStSaut := ReadTokenSt(lStRetour) ;
 lStRup  := ReadTokenSt(lStRetour) ;

 lTOBDup := TOB.Create('' , nil , -1 ) ;

 try

 lTOBDup.Dupliquer(ATobFListe,true,true) ;
 lTOBDup.Detail[0].AddChampSup('SOLDECOMPTE', true) ;
 lTOBDup.Detail[0].AddChampSup('REG1'       , true ) ;
 lTOBDup.Detail[0].AddChampSup('SAUTPAGE'   , true ) ;
 lTOBDup.Detail[0].AddChampSup('DATEC'      , true ) ;
 lTOBDup.Detail[0].AddChampSup('DATEC1'     , true ) ;

 for i := 0 to lTOBDup.Detail.Count - 1 do
  begin
   lTOB := lTOBDup.Detail[i] ;
   if length(trim(lStRup)) <> 0 then
    lTOB.PutValue('REG1',  lTOB.GetValue(lStRup) ) ;
   lTOB.PutValue('SAUTPAGE',  lStSaut ) ;
   lTOB.PutValue('DATEC',  GetControlText('DATEC') ) ;
   lTOB.PutValue('DATEC1',  DateToStr(StrToDate(GetControlText('DATEC')) + 1 ) ) ;
   FInfo.LoadCompte(lTOB.GetValue('E_GENERAL')) ;
   lTOB.PutValue('SOLDECOMPTE', FInfo.GetValue('G_TOTDEBE')-FInfo.GetValue('G_TOTCREE') ) ;
  end ; // for

 if trim(UpperCase(lStSaut)) = 'X' then
  lStCrit := 'SAUTPAGE=X' + '`'
   else
    lStCrit := '' ;

 lTOBDup.Detail.Sort('G_GENERAL;' + lStRup + ';E_DATECOMPTABLE' );
 LanceEtatTOB('E','CUT','CU1',lTOBDup,true,false,false,nil,'','Justification de comptes',false,0,lStCrit) ;

 finally
  lTOBDup.Free ;
 end ;

end;


{ TOF_CPQRCUTOFF }

procedure TOF_CPCUTOFFPARAM.OnArgument(S: String);
begin
 SetControlText('SAUT','Aucune');
 TToolbarButton97(GetControl('BValider')).Onclick := BValiderClick;
 inherited;
end;

procedure TOF_CPCUTOFFPARAM.OnClose;
begin
 //TFVierge(Ecran).Retour := GetControlText('SAUT') + ' ; ' + GetControlText('RUPTURE') ;
 inherited;
end;

procedure TOF_CPCUTOFFPARAM.OnLoad ;
begin
 Inherited ;
end ;

procedure TOF_CPCUTOFFPARAM.BValiderClick(Sender: TObject);
begin
 TFVierge(Ecran).retour := GetControlText('SAUT') + ' ; ' + GetControlText('RUPTURE') ;
end;


Initialization
  registerclasses ( [ TOF_CPMULCUTOFF ] ) ;
  registerclasses ( [ TOF_CPCUTOFFPARAM ] ) ;

end.


