{$A+,B-,C+,D+,E-,F-,G+,H+,I+,J+,K-,L+,M-,N+,O+,P+,Q-,R-,S-,T-,U-,V+,W-,X+,Y+,Z1}
{$MINSTACKSIZE $00004000}
{$MAXSTACKSIZE $00100000}
{$IMAGEBASE $00400000}
{$APPTYPE GUI}

unit CPGESTIONCREANCE ;

interface

uses
 {$IFNDEF EAGLCLIENT}
  db,
 {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
 {$ENDIF}
 UTOB ,
 Controls ,
 Classes ,
 ULibEcriture ,
 UTOM,
 UTOF,
 HCtrls ,
 uTofViergeMul ;   // Fiche Ancetre Vierge MUL

type

  TTypeCreance = ( GCDDeb, GCDFin , GCDNew , GCDRegl , GCDPerte , GCDProv ) ;

  TOF_CPMULGCD = Class (TOF_ViergeMul)
   private
    FTOBEcr       : TOB ;
    FInfo         : TInfoEcriture ;
    FColCount     : integer ;
    FRappro       : boolean ;
    procedure OnPopUpPopUpTraitement(Sender: TObject) ;
    procedure OnClickTaux ( Sender : TObject) ;
    procedure OnClickCreance ( Sender : TObject) ;
    procedure OnClickPerte ( Sender : TObject) ;
    procedure OnClickPOP6 ( Sender : TObject ) ;
    procedure OnClickPOP3 ( Sender : TObject ) ;
    procedure OnClickPOP8 ( Sender : TObject ) ;
    procedure OnClickPOP9 ( Sender : TObject ) ;
    procedure OnClickPOP10 ( Sender : TObject ) ;
    procedure GCDCompta ;
    procedure FListeOnDblClick ( Sender : TObject ) ;
    procedure InitPopUp ;
    function  ContruirePieceFin : integer ;
    procedure TransactionPieceFin ;
   protected
    procedure RemplitATobFListe ; override ;
    procedure InitControl ; override ;
    function  AjouteATobFListe( vTob : Tob) : Boolean; override ;
    function  AfterLoad  : boolean ; override ;
   public
    procedure OnArgument              (S : String ) ; override ;
    procedure OnClose ; override ;
    procedure OnClickBAgrandir     (Sender : Tobject); override ;
    procedure OnClickBReduire      (Sender : Tobject); override ;
    procedure OnKeyDownFListe      (Sender : TObject; var Key : Word; Shift : TShiftState ); override ;
    procedure OnKeyDownEcran( Sender : TObject; var Key : Word; Shift : TShiftState ) ; override ;
    procedure OnRowEnterFListe (Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean); override ;

    procedure OnDelete ; override;
    procedure OnNew ; override;
    procedure OnPopUpPopF11 ( Sender : TObject) ; override ;

  end ;

  TZGCD = Class(TObjetCompta)
  private
   PTOBEcr                   : TOB ;
   FGOP                      : TOB ;
   PTOBALettrer              : TOB ;
   FBoALettrer               : boolean ;
   FBoAGenerer               : boolean ;
   FBoFaireEcr               : boolean ;
   FTTC                      : double ;
   FTVA                      : double ;
   FHT                       : double ;
   FTOBCumul                 : TOB ;
   FAUXILIAIRE               : string ;
   FAUXILIAIRECONTRE         : string ;
   FGENERAL                  : string ;
   FGENERALCREANCE           : string ;
   FGENERALCONTRE            : string ;
   FJOURNAL                  : string ;
   FNATUREPIECE              : string ;
   FLIBELLE                  : string ;
   FLIBELLEECR               : string ;
   FCODETAUX                 : string ;
   FDATECOMPTABLE            : TDateTime ;
   FCATEGORIE                : TTypeCreance ;
   FPROVISION                : double ;
   FTAUX                     : double ;
   FTAUXBASE                 : double ;
   FCOMPTETVA                : string ;
   FBoForceTaux              : boolean ;
   function  GetCumulTTC : double ;
   function  GetCumulHT  : double ;
   function  GetSolde    : double ;
   procedure SetDate ( Value : TDateTime ) ;
   procedure SetCategorie ( Value : TTypeCreance ) ;
   procedure SetTTC ( Value : double ) ;
   procedure SetTaux ( Value : double ) ;
   procedure SetProvision ( Value : double ) ;
   procedure ConstruirePieceCreance ;
   procedure ConstruirePiecePerte ;
   procedure ConstruirePieceLigneUn ;
   procedure MajCumuls ;
   procedure RAZCumul ;
  public

   constructor Create( vInfoEcr : TInfoEcriture ) ; override ;
   destructor  Destroy ; override ;

   function    ConstruirePiece : boolean ;
   function    EnregistreCreance : boolean ;
   function    EnregistrePiece : boolean ;
   function    Lettrage : boolean ;
   function    IsValidTTC : integer ;
   function    IsValidDate : integer ;
   function    IsValidAux  : integer ;
   function    IsValid : integer ;
   procedure   Load ( Value : string ) ;
   procedure   LoadFromTOB ( Value : TOB ) ;
   function    GetLibelle : string ;
   procedure   SetLibelle ( Value : string ) ;

   property CATEGORIE          : TTypeCreance read FCATEGORIE              write SetCATEGORIE ;
   property TTC                : double       read FTTC                    write SetTTC ;
   property TVA                : double       read FTVA                    write FTVA ;
   property HT                 : double       read FHT                     write FHT ;
   property AUXILIAIRE         : string       read FAUXILIAIRE ;
   property GENERAL            : string       read FGENERAL                write FGENERAL ;
   property JOURNAL            : string       read FJOURNAL                write FJOURNAL ;
   property NATUREPIECE        : string       read FNATUREPIECE            write FNATUREPIECE ;
   property LIBELLE            : string       read GetLIBELLE              write SetLIBELLE ;
   property DATECOMPTABLE      : TDateTime    read FDATECOMPTABLE          write SetDate ;
   property CODETAUX           : string       read FCODETAUX               write FCODETAUX ;
   property TAUX               : double       read FTAUX                   write SetTaux ;
   property TAUXBASE           : double       read FTAUXBASE               write FTAUXBASE ;
   property PROVISION          : double       read FPROVISION              write SetProvision ;
   property GENERALCONTRE      : string       read FGENERALCONTRE          write FGENERALCONTRE ;
   property AUXILIAIRECONTRE   : string       read FAUXILIAIRECONTRE       write FAUXILIAIRECONTRE ;

   property CumulTTC           : double       read GetCumulTTC ;
   property CumulHT            : double       read GetCumulHT ;
   property Solde              : double       read GetSolde ;
   property ALettrer           : boolean      read FBoALettrer ;
   property AGenerer           : boolean      read FBoAGenerer ;
   property FaireEcr           : boolean      read FBoFaireEcr             write FBoFaireEcr ;
   property Ecr                : TOB          read FGOP ;


  end ;

 TOF_CPGCDOPERATION = Class (TOF)
  private
   FGCD             : TZGCD ;
   FInfo            : TInfoEcriture ;
   FMessCompta      : TMessageCompta ;        // affichage des messages
   FOldKeyDown      : TKeyEvent ;
   procedure OnErrorTOB ( sender : TObject ; Error : TRecError) ;
   procedure TransactionCreance ;
   procedure TransactionLettrage ;
   procedure TransactionPiece ;
   procedure OnExitDate ( Sender : TObject ) ;
   procedure OnExitTTC ( Sender : TObject ) ;
   procedure OnExitTVA ( Sender : TObject ) ;
   procedure OnExitAux ( Sender : TObject ) ;
   procedure OnExitLibelle ( Sender : TObject ) ;
   procedure OnExitProvision ( Sender : TObject ) ;
   procedure OnExitTaux ( Sender : TObject ) ;
   procedure OnChangeCodeTaux ( Sender : TObject ) ;
   procedure OnClickBPARAM ( Sender : TObject ) ;
   procedure OnKeyDownEcran( Sender : TObject; var Key : Word; Shift : TShiftState ) ;
   procedure InitControl ;
   procedure CalculSolde ;
  public
   procedure OnArgument (S : String ) ; override ;
   procedure OnClose                  ; override ;
   procedure OnUpdate                 ; override ;
  end ;

  TOF_CPGCDLISTEOPE = Class (TOF)
  private
   FListe : THGrid ;
   FAux   : string ;
   FRow   : integer ;
   FTOB   : TOB ;
   procedure RemplirFListe ;
   procedure MajCumuls ;
  public
   procedure OnArgument (S : String ) ; override ;
   procedure OnClose                  ; override ;
   procedure OnDelete                 ; override ;
  end ;


 TOM_CPGCDTAUX = Class (TOM)
  procedure OnClickBValider( Sender : TObject ) ; 
  procedure OnNewRecord                ; override ;
  procedure OnDeleteRecord             ; override ;
  procedure OnUpdateRecord             ; override ;
  procedure OnLoadRecord               ; override ;
  procedure OnChangeField ( F: TField) ; override ;
  procedure OnArgument ( S: String )   ; override ;
  procedure OnClose                    ; override ;
  end ;

function CPLanceFiche_CPMULGCD( vStParam : string = '' ) : string ;
function CPLanceFiche_GCDOperation( vStLequel : string = ''; vStParam : string = '') : string ;
procedure CPLanceEtatGCD ;

Implementation

uses
 {$IFNDEF EAGLCLIENT}
 FichList ,
 {$ELSE}
 eFichList ,
 {$ENDIF}
 {$IFDEF MODENT1}
 CPTypeCons,
 {$ENDIF MODENT1}
 Vierge ,
 SysUtils ,
 Windows ,
 ULibWindows ,
 ULibExercice ,
 Ent1 ,
 AglInit ,
 hmsgbox ,
 hEnt1 ,
 LettUtil,
 ULibAnalytique ,
 ParamSoc,
 Saiscomm,
 Menus,           // PopUpMenu
 HTB97,
 Saisie,
 SaisBor ,
 ZCompte ,
{$IFDEF EAGLCLIENT}
 UtileAGL,      // LanceEtat
 MainEagl; // AglLanceFiche
{$ELSE}
 EdtREtat,      // LanceEtat
 FE_Main; // AglLanceFiche
{$ENDIF}

const
 cFI_TABLE    = 'CPGCDCUMULS' ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 17/07/2007
Modifié le ... :   /  /
Description .. : - LG - 17/07/2007 - FB 21072 - coorection du msg d'erreur
Mots clefs ... :
*****************************************************************}
function _ControleParamSoc : boolean ;
var
 lCompte   : TZCompte ;
 lStCompte : string ;
begin

 result := false ;

 lCompte   := TZCompte.Create('') ;
 lStCompte := GetParamSocSecur('SO_CPGCDGENERAL','') ;

 try

  if lCompte.GetCompte(lStCompte) = - 1 then
  begin
   PGIError('Le compte de gestion des créances n''est pas défini !') ;
   exit ;
  end ;


 if trim(GetParamSocSecur('SO_CPGCDJAL','')) = '' then
  begin
   PGIError('Le journal de génération des créances n''est pas défini !') ;
   exit ;
  end ;

 result := true ;

 finally
  lCompte.Free ;
 end ;

end ;


function CPLanceFiche_CPMULGCD( vStParam : string = '' ) : string ;
begin
  result := AGLLanceFiche('CP', 'CPMULGCD', '', '', vStParam ) ;
end ;

function CPLanceFiche_GCDOperation(vStLequel : string = '' ; vStParam : string = '') : string;
begin

 if not _ControleParamSoc then exit ;

 if ( vStParam = 'CONSULT' ) and ( TheData = nil ) then
  begin
   PGIError('Les données de départ sont manquantes !') ;
   exit ;
  end ;

 AGLLanceFiche('CP', 'CPGCDOPERATION', '', vStLequel , vStParam );

end ;


procedure CPLanceEtatGCD ;
begin
 LanceEtat('E','CGC','CG1',true,false,false,nil,'','Gestion des créances douteuses',false) ;
end ;


function _ProchainNumero : integer;
begin
 // on calcul le prochain numero et on le met à jour en meme temps
 result := CCalculProchainNumeroSouche ('CPT','GCD');
 if result = -1 then
  begin // creation de la souche
   ExecuteSQL('insert into SOUCHE( '                 +
              'SH_TYPE, '                            +
              'SH_SOUCHE, '                          +
              'SH_LIBELLE, '                         +
              'SH_ABREGE, '                          +
              'SH_NATUREPIECE, '                     +
              'SH_NUMDEPART, '                       +
              'SH_JOURNAL, '                         +
              'SH_MASQUENUM, '                       +
              'SH_SOCIETE, '                         +
              'SH_DATEDEBUT, '                       +
              'SH_DATEFIN, '                         +
              'SH_FERME, '                           +
              'SH_ANALYTIQUE, '                      +
              'SH_NATUREPIECEG, '                    +
              'SH_SIMULATION, '                      +
              'SH_NUMDEPARTS, '                      +
              'SH_NUMDEPARTP, '                      +
              'SH_SOUCHEEXO, '                       +
              'SH_RESERVEWEB ) '                     +
              ' values ('                            +
              '"CPT", '                              +
              '"GCD", '                              +
              '"Compteur Créance douteuse", '        +
              '"GCD", '                              +
              'NULL, '                               +
              '1, '                                  +
              'NULL, '                               +
              'NULL, '                               +
              '001, '                                +
              '"'+UsDateTime(iDate1900)+'", '        +
              '"'+UsDateTime(iDate2099)+'", '        +
              '"-", '                                +
              '"-", '                                +
              '"", '                                 +
              '"-", '                                +
              '1, '                                  +
              '1, '                                  +
              '"-", '                                +
              '"-" ) ');
   result := 1;
  end; // if

end;

procedure _OuvrirPiece( vTOB : TOB ) ;
var
 lTOB : TOB ;
 lQ   : TQuery ;
begin

 if ( vTOB = nil ) or ( vTOB.Detail = nil ) or ( VTOB.Detail.Count = 0 ) then exit ;

 if ( V_PGI.IoError = oeOk ) and ( PGIAsk('Traitement terminé, voulez vous voir les écritures générées ?') = mrYes ) then
   begin
    lTOB := vTOB.Detail[0] ;
    lQ   := OpenSql('SELECT E_DATECOMPTABLE, E_JOURNAL, E_NUMEROPIECE, ' +
             'E_NUMLIGNE FROM ECRITURE WHERE ' +
             'E_EXERCICE = "' + lTOB.GetString('E_EXERCICE') + '" AND ' +
             'E_DATECOMPTABLE >= "' + USDateTime(lTOB.GetValue('E_DATECOMPTABLE')) + '" AND ' +
             'E_DATECOMPTABLE <= "' + USDateTime(lTOB.GetValue('E_DATECOMPTABLE')) + '" AND ' +
             'E_JOURNAL = "' + lTOB.GetString('E_JOURNAL') + '" AND ' +
             'E_NUMEROPIECE = ' + VariantToSql(lTOB.GetValue('E_NUMEROPIECE')) + ' AND ' +
             'E_NUMLIGNE = 1 AND E_NUMECHE <= 1 AND E_QUALIFPIECE = "N" ' +
             'ORDER BY E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE, E_QUALIFPIECE', True);

    if (lTOB.GetString('E_MODESAISIE') <> '-') and (lTOB.GetString('E_MODESAISIE') <> '') then
     LanceSaisieFolio(lQ, taModif,true)
      else
       TrouveEtLanceSaisie(lQ, taModif, 'N') ;
    Ferme(lQ) ;
   end ;
end ;



{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 30/08/2007
Modifié le ... :   /  /    
Description .. : - LG - 30/08/2007 - Fb 21279 - correction du arrondi ds les 
Suite ........ : calculs
Mots clefs ... : 
*****************************************************************}
function TOF_CPMULGCD.AjouteATobFListe( vTob : Tob ) : Boolean;
var
 lRdVal       : double ;
 lRdVariation : double ;
begin
 lRdVal := vTOB.GetValue('CGC_SOLDE') + vTOB.GetValue('CGC_CREANCE') - vTOB.GetValue('CGC_REGLEMENT') - vTOB.GetValue('CGC_PERTE') ;
 vTob.AddChampSupValeur('FINEXO'     , lRdVal ) ;
 vTob.AddChampSupValeur('DOTATIONS'  , 0 ) ;
 vTob.AddChampSupValeur('REPRISES'   , 0 ) ;
 lRdVal       := ( vTOB.GetValue('CGC_HT') * vTOB.GetValue('CGC_TAUX') ) / 100 ;
 vTob.AddChampSupValeur('FINEXOP'    , Arrondi(lRdVal,2) ) ;
 lRdVariation := Arrondi(lRdVal - vTOB.GetValue('CGC_PROVISION') , 2 ) ;
 if lRdVariation > 0 then
  vTob.PutValue('DOTATIONS' , lRdVariation )
   else
    vTob.PutValue('REPRISES' , lRdVariation * (-1) ) ;

 result := true ;

end;


procedure TOF_CPMULGCD.OnArgument(S: String);
begin

 FFI_Table       := cFI_TABLE;
 FStListeParam   := 'CPGCDCUMULS';
 FRappro         := false ;
 AboFetchLesTous := true ;
 FInfo           := TInfoEcriture.Create();

 inherited ;

 FListe.OnDblClick := FListeOnDblClick ;
 FColCount         := FListe.ColCount ;
 BAgrandir.Hint    := TraduireMemoire('Rapprochement GCD/Comptabilité') ;
 BReduire.Hint     := TraduireMemoire('Tableau justificatifs') ;

 if S = 'GCDCOMPTA' then
  begin
   RefreshPclPge ;
   GCDCompta ;
  end
   else
    InitPopUp ;

end;

procedure TOF_CPMULGCD.OnClose;
begin
 FreeAndNil(FInfo) ;
 inherited ;
end;

procedure TOF_CPMULGCD.OnKeyDownEcran( Sender : TObject; var Key : Word; Shift : TShiftState ) ;
begin

 if csDestroying in Ecran.ComponentState then Exit ;


 case Key of
  //VK_Space  : if not FRappro then
  //             Inherited OnKeyDownEcran(Sender,Key,Shift) ;
  VK_DELETE : if  Shift=[ssCtrl] then begin Key:=0 ; OnDelete ; end ;
  VK_F12    : begin Key := 0 ; exit ; end ;
  78 : if Shift = [ssCtrl] then // ctrl + N
        begin
         key := 0 ;
         OnClickCreance (nil) ;
        end ;
 else
  inherited ;
//  FOldKeyDown(Sender,Key,Shift) ;
 end ;

end ;

procedure TOF_CPMULGCD.OnKeyDownFListe(Sender: TObject; var Key: Word; Shift: TShiftState) ;
begin

 if csDestroying in Ecran.ComponentState then Exit ;

  case Key of
   VK_Space  : if not FRappro then
                Inherited OnKeyDownFListe(Sender, Key , Shift ) ;
  end ; // case

end ;


procedure TOF_CPMULGCD.RemplitATobFListe;
begin

 inherited;

 AStSqlTobFListe := 'select * from CPGCDCUMULS' ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 20/08/2007
Modifié le ... : 29/08/2007
Description .. : - LG - FB 21230 - remise a jour de la tablette apres une 
Suite ........ : modif
Suite ........ : - FB 21271 - correction pour les auxi trop court
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPMULGCD.OnClickTaux ( Sender : TObject) ;
var
 lStRetour : string ;
 lStVal    : string ;
 lTOB      : TOB ;
 i         : integer ;
begin

 if csDestroying in Ecran.ComponentState then Exit ;

 lStRetour := AGLLanceFiche('CP', 'CPGCDTAUX','','', 'ACTION=CONSULT' ) ;

 if ( lStRetour <> '' ) then
  begin

   lStVal := ReadTokenSt(lStRetour) ;

   ExecuteSQL('update CPGCDCUMULS set CGC_TAUX=' + VariantToSql(Valeur (lStRetour)) +
              ' where CGC_CODETAUX="' + lStVal +'" ' ) ;

   if ABoMultiSelected then
    begin

     for i := 1 to FListe.RowCount - 1 do
      if FListe.IsSelected(i) then
       begin
        lTob := GetO(FListe,i) ;
        if lTob = nil then continue ;
        ExecuteSQL('update CPGCDCUMULS set CGC_CODETAUX="' + lStVal +'" , ' +
               ' CGC_TAUX=' + VariantToSql(Valeur (lStRetour)) +
               ' where CGC_AUXILIAIRE="' + lTOB.GetValue('CGC_AUXILIAIRE') + '" ' ) ;

        ExecuteSQL('update CPGCDOPERATION set GOP_CODETAUX="' + lStVal +'" ' +
                   ' WHERE GOP_NUMERO IN (SELECT CND_NUMEROSEQ FROM CNTVDETAIL WHERE CND_PERM="1" AND CND_AUXILIAIRE="' + lTob.GetString('CGC_AUXILIAIRE') + '") ' ) ;

      //  ExecuteSQL('update CPGCDOPERATION set GOP_CODETAUX="' + lStVal +'" ' +
      //        ' where GOP_AUXILIAIRE="' + Copy(lTOB.GetValue('CGC_AUXILIAIRE'),1,6) + '" AND  ' ) ;

      end; // if
    end
     else
      begin
        lTob := GetO(FListe) ;      
        if lTob = nil then exit ;
        ExecuteSQL('update CPGCDCUMULS set CGC_CODETAUX="' + lStVal +'" , ' +
                ' CGC_TAUX=' + VariantToSql(Valeur (lStRetour)) +
                ' where CGC_AUXILIAIRE="' + lTOB.GetValue('CGC_AUXILIAIRE') + '" ' ) ;
        ExecuteSQL('update CPGCDOPERATION set GOP_CODETAUX="' + lStVal +'" ' +
                   ' WHERE GOP_NUMERO IN (SELECT CND_NUMEROSEQ FROM CNTVDETAIL WHERE CND_PERM="1" AND CND_AUXILIAIRE="' + lTob.GetString('CGC_AUXILIAIRE') + '") ' ) ;
      end ;

   RefreshPclPge ;
  end ;

 RemplirListe('CPGCDTAUX','');

end ;

procedure TOF_CPMULGCD.OnClickCreance ( Sender : TObject) ;
var
 lTOB   : TOB ;
 lStAux : string ;
begin

 if csDestroying in Ecran.ComponentState then Exit ;

 lStAux := '' ;
 lTOB   := GetO(FListe) ;

 if lTOB <> nil then
  lStAux := lTOB.GetValue('CGC_AUXILIAIRE') ;

 CPLanceFiche_GCDOperation('' ,'TAB;CREANCE;' + lStAux ) ;
 RefreshPclPge ;

end ;

procedure TOF_CPMULGCD.OnClickPOP6 ( Sender : TObject) ;
begin
 if csDestroying in Ecran.ComponentState then Exit ;
 OnDelete ;
end ;

procedure TOF_CPMULGCD.OnClickPOP3 ( Sender : TObject) ;
var
 lTOB   : TOB ;
 lStAux : string ;
begin

 if csDestroying in Ecran.ComponentState then Exit ;

 lStAux := '' ;
 lTOB   := GetO(FListe) ;

 if lTOB <> nil then
  lStAux := lTOB.GetValue('CGC_AUXILIAIRE') ;

 CPLanceFiche_GCDOperation('' ,'TAB;PROVISION;' + lStAux ) ;

 RefreshPclPge ;

end ;


procedure TOF_CPMULGCD.OnClickPerte ( Sender : TObject) ;
var
 lTOB   : TOB ;
 lStAux : string ;
begin

 if csDestroying in Ecran.ComponentState then Exit ;

 lStAux := '' ;
 lTOB   := GetO(FListe) ;

 if lTOB <> nil then
  lStAux := lTOB.GetValue('CGC_AUXILIAIRE') ;

 CPLanceFiche_GCDOperation('','TAB;PERTE;' + lStAux ) ;

 RefreshPclPge ;

end ;


procedure TOF_CPMULGCD.OnPopUpPopUpTraitement(Sender: TObject);
var
 i                : integer ;
 lTOB             : TOB ;
 lPopUpTraitement : TPopUpMenu ;
begin

 if csDestroying in Ecran.ComponentState then Exit ;

 lTOB := GetO(FListe) ;

 lPopUpTraitement := TPopUpMenu(GetControl('PopUpTraitement', True)) ;

 for i := 0 to lPopUpTraitement.Items.Count - 1 do
  begin

 {   if lPopUpTraitement.Items[i].Name = 'POP4' then
     begin
      lPopUpTraitement.Items[i].Enabled := lTOB <> nil ;
      Continue;
     end;  }

    if lPopUpTraitement.Items[i].Name = 'POP6' then
     begin
      lPopUpTraitement.Items[i].Enabled := lTOB <> nil ;
      Continue;
     end;


    if lPopUpTraitement.Items[i].Name = 'POP2' then
     begin
      lPopUpTraitement.Items[i].Enabled := lTOB <> nil ;
      Continue;
     end ;

    if lPopUpTraitement.Items[i].Name = 'POP1' then
     begin
      lPopUpTraitement.Items[i].Enabled := not FRappro ;
      Continue;
     end;

    if lPopUpTraitement.Items[i].Name = 'POP3' then
     begin
      lPopUpTraitement.Items[i].Enabled := not FRappro and ( lTOB <> nil ) and ( lTOB.GetValue('CGC_SOLDE') > 0 ) ;
      Continue;
     end;

    if lPopUpTraitement.Items[i].Name = 'POP4' then
     begin
      lPopUpTraitement.Items[i].Enabled := false ; // desactiver pour l'instant
      Continue;
     end;

  end ; // for
end;

procedure TOF_CPMULGCD.OnPopUpPopF11(Sender: TObject);
begin
 OnPopUpPopUpTraitement(sender) ;
 inherited; // Ajoute les élements visibles des menus de Ecran dans PopF11
end;

procedure TOF_CPMULGCD.InitPopUp ;
var
 lPopUpTraitement : TPopUpMenu ;
 lPopUpAutre      : TPopUpMenu ;
 i                : integer ;
begin

 if csDestroying in Ecran.ComponentState then Exit ;

 lPopUpTraitement := TPopUpMenu(GetControl('PopUpTraitement', True)) ;

 for i := 0 to lPopUpTraitement.Items.Count - 1 do
   begin

    if lPopUpTraitement.Items[i].Name = 'POP7' then
     begin
      lPopUpTraitement.Items[i].OnClick := OnClickTaux ;
      Continue;
     end;

    if lPopUpTraitement.Items[i].Name = 'POP2' then
     begin
      lPopUpTraitement.Items[i].OnClick := OnClickPerte ;
      Continue;
     end;

    if lPopUpTraitement.Items[i].Name = 'POP3' then
     begin // menu suppression
      lPopUpTraitement.Items[i].OnClick := OnClickPOP3 ;
      Continue;
     end;

    if lPopUpTraitement.Items[i].Name = 'POP8' then
     begin // menu suppression
      lPopUpTraitement.Items[i].OnClick := OnClickPOP8 ;
      Continue;
     end;

    if lPopUpTraitement.Items[i].Name = 'POP9' then
     begin // menu suppression
      lPopUpTraitement.Items[i].OnClick := OnClickPOP9 ;
      Continue;
     end;

    if lPopUpTraitement.Items[i].Name = 'POP10' then
     begin // menu suppression
      lPopUpTraitement.Items[i].OnClick := OnClickPOP10 ;
      Continue;
     end;

   end ;

  lPopUpAutre := TPopUpMenu(GetControl('POPUPAUTRE', True)) ;

  for i := 0 to lPopUpAutre.Items.Count - 1 do
   begin

    if lPopUpAutre.Items[i].Name = 'POP6' then
     begin // menu suppression
      lPopUpAutre.Items[i].OnClick := OnClickPOP6 ;
      Continue;
     end;

    if lPopUpAutre.Items[i].Name = 'POP1' then
     begin
      lPopUpAutre.Items[i].OnClick := OnClickCreance ;
      Continue;
     end;

   end ; // for


end ;

procedure TOF_CPMULGCD.OnNew ;
begin

 CPLanceFiche_GCDOperation('','TAB;CREANCE;') ;

 RefreshPclPge ;
 
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 29/08/2007
Modifié le ... :   /  /    
Description .. : - FB 21271 - correction pour les auxi trop court
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPMULGCD.OnDelete ;
var
 lTobListe : TOB ;
 i         : integer ;
begin

 if csDestroying in Ecran.ComponentState then Exit ;

 if PGIAsk('Voulez-vous supprimer les créances selectionnées ?') = mrNo then exit ;

 if ABoMultiSelected then
  begin

    for i := 1 to FListe.RowCount - 1 do
     if FListe.IsSelected(i) then
      begin
       lTobListe := GetO(FListe,i) ;
       if lTobListe = nil then continue ;
       ExecuteSQL('DELETE CPGCDCUMULS WHERE CGC_AUXILIAIRE="' + lTobListe.GetString('CGC_AUXILIAIRE') + '" ' ) ;
       ExecuteSQL('DELETE CPGCDOPERATION WHERE GOP_NUMERO IN (SELECT CND_NUMEROSEQ FROM CNTVDETAIL WHERE CND_AUXILIAIRE="' + lTobListe.GetString('CGC_AUXILIAIRE') + '") ' ) ;
       ExecuteSQL('DELETE CNTVDETAIL WHERE CND_AUXILIAIRE="' + lTobListe.GetString('CGC_AUXILIAIRE') + '" ' ) ;
      end; // if
  end
   else
    begin
      lTobListe := GetO(FListe) ;
      if lTobListe = nil then exit ;
      ExecuteSQL('DELETE CPGCDCUMULS WHERE CGC_AUXILIAIRE="' + lTobListe.GetString('CGC_AUXILIAIRE') + '" ' ) ;
      ExecuteSQL('DELETE CPGCDOPERATION WHERE GOP_NUMERO IN (SELECT CND_NUMEROSEQ FROM CNTVDETAIL WHERE CND_AUXILIAIRE="' + lTobListe.GetString('CGC_AUXILIAIRE') + '") ' ) ;
      ExecuteSQL('DELETE CNTVDETAIL WHERE CND_AUXILIAIRE="' + lTobListe.GetString('CGC_AUXILIAIRE') + '" ' ) ;
    end ;

  RefreshPclPge ;

  inherited ;

end ;

procedure TOF_CPMULGCD.FListeOnDblClick ( Sender : TObject ) ;
var
 lTOB : TOB ;
begin

 if csDestroying in Ecran.ComponentState then Exit ;

 lTOB := GetO(FListe) ;

 if lTOB <> nil then
  begin
   AGLLanceFiche('CP', 'CPGCDLISTEOP', '', '', lTOB.GetString('CGC_AUXILIAIRE') ) ;
   RefreshPclPge ;
  end ;
end ;

function TOF_CPMULGCD.ContruirePieceFin : integer ;
var
 lStReprise    : string ;
 lStDotation   : string ;
 lStProvision  : string ;
 lRdDotation   : double ;
 lRdReprise    : double ;
 lRdSolde      : double ;
 lStJal        : string ;
 lTOB          : TOB ;
 lDtDate       : TDateTime ;
begin

 result := RC_BADWRITE ;

 if existeSQL('select E_EXERCICE from ecriture where E_QUALIFORIGINE="GCF" and E_EXERCICE="' + ctxExercice.EnCours.Code + '" ' ) then
  begin
   if PGIAsk('Il existe déjà des écritures, voulez vous quand même les regénérer ?') = mrNo then
    exit ;
  end ;

 lStReprise    := GetParamSocSecur('SO_CPGCDREPRISE' , '' ) ;
 lStDotation   := GetParamSocSecur('SO_CPGCDDOTPROV' , '' ) ;
 lStProvision  := GetParamSocSecur('SO_CPGCDPROVISION' , '' ) ;
 lStJal        := GetParamSocSecur('SO_CPGCDJAL' , '') ;
 lDtDate       := ctxExercice.EnCours.Fin ;

 FInfo.Devise.Load([V_PGI.DevisePivot]) ;

 FInfo.LoadJournal(lStJal) ;

 lRdDotation := Arrondi( GetCumulChamps('DOTATIONS') , 2 ) ;
 lRdReprise  := Arrondi ( GetCumulChamps('REPRISES') , 2 ) ;

 if ( lRdDotation + lRdReprise ) = 0 then
  begin
   PGIInfo('Il n''y a pas de dotations et de reprises à comptabiliser') ;
   result := RC_BADWRITE ;
   exit ;
  end ;

 lRdSolde := arrondi(lRdDotation - lRdReprise , 2 ) ;

 // 1er ligne
 if lRdDotation <> 0 then
  begin
   FInfo.LoadCompte(lStDotation) ;
   lTOB    := TOB.Create('ECRITURE' , FTOBEcr , - 1 ) ;
   CPutDefautEcr(lTOB) ;
   lTOB.PutValue('E_NUMLIGNE'       , -1 ) ;
   lTOB.PutValue('E_NUMEROPIECE'    , -1 ) ;
   lTOB.PutValue('E_NUMGROUPEECR'   , 0 ) ;
   lTOB.PutValue('E_LIBELLE'        , 'INVENTAIRE GCD' ) ;
   lTOB.PutValue('E_GENERAL'        , lStDotation ) ;
   lTOB.PutValue('E_JOURNAL'        , lStJal ) ;
   lTOB.PutValue('E_MODESAISIE'     , FInfo.GetString('J_MODESAISIE') ) ;
   lTOB.PutValue('E_IO'             , 'X') ;
   lTOB.PutValue('E_QUALIFORIGINE'  , 'GCF' ) ;
   CRemplirDateComptable(lTOB , lDtDate ) ;

   CSetMontants( lTOB , lRdDotation , 0 , FInfo.Devise.Dev ) ;
   CVentilerTOB(lTOB,FInfo) ;

   if FInfo.Compte.IsLettrable then
    CRemplirInfoLettrage(lTOB)
     else
      CSupprimerInfoLettrage(lTOB) ;

   if FInfo.Compte.IsPointable then
    CRemplirInfoPointage(lTOB) ;

   CGetTVA(lTOB,FInfo) ;
   CGetEch(lTOB,FInfo) ;
  end ;

 // 2eme ligne
 FInfo.LoadCompte(lStReprise) ;

 if lRdReprise <> 0 then
  begin
   lTOB    := TOB.Create('ECRITURE' , FTOBEcr , - 1 ) ;
   CPutDefautEcr(lTOB) ;
   lTOB.PutValue('E_NUMLIGNE'       , -1 ) ;
   lTOB.PutValue('E_NUMEROPIECE'    , -1 ) ;
   lTOB.PutValue('E_NUMGROUPEECR'   , 0 ) ;
   lTOB.PutValue('E_LIBELLE'        , 'INVENTAIRE GCD' ) ;
   lTOB.PutValue('E_GENERAL'        , lStReprise ) ;
   lTOB.PutValue('E_JOURNAL'        , lStJal ) ;
   lTOB.PutValue('E_MODESAISIE'     , FInfo.GetString('J_MODESAISIE') ) ;
   lTOB.PutValue('E_IO'             , 'X') ;
   lTOB.PutValue('E_QUALIFORIGINE'  , 'GCF' ) ;
   CRemplirDateComptable(lTOB , lDtDate ) ;

   CSetMontants( lTOB , 0 , lRdReprise , FInfo.Devise.Dev ) ;
   CVentilerTOB(lTOB,FInfo) ;

   if FInfo.Compte.IsLettrable then
    CRemplirInfoLettrage(lTOB)
     else
      CSupprimerInfoLettrage(lTOB) ;

   if FInfo.Compte.IsPointable then
    CRemplirInfoPointage(lTOB) ;

   CGetTVA(lTOB,FInfo) ;
   CGetEch(lTOB,FInfo) ;
  end ;

 //3eme ligne
 if lRdSolde <> 0 then
  begin
   FInfo.LoadCompte(lStProvision) ;
   lTOB    := TOB.Create('ECRITURE' , FTOBECR , - 1 ) ;
   CPutDefautEcr(lTOB) ;
   lTOB.PutValue('E_NUMLIGNE'       , -1 ) ;
   lTOB.PutValue('E_NUMEROPIECE'    , -1 ) ;
   lTOB.PutValue('E_NUMGROUPEECR'   , 0 ) ;
   lTOB.PutValue('E_LIBELLE'        , 'INVENTAIRE GCD' ) ;
   lTOB.PutValue('E_GENERAL'        , lStProvision ) ;
   lTOB.PutValue('E_JOURNAL'        , lStJal ) ;
   lTOB.PutValue('E_MODESAISIE'     , FInfo.GetString('J_MODESAISIE') ) ;
   lTOB.PutValue('E_IO'             , 'X') ;
   lTOB.PutValue('E_QUALIFORIGINE'  , 'GCF' ) ;
   CRemplirDateComptable(lTOB , lDtDate ) ;

   if lRdSolde > 0 then
    CSetMontants( lTOB , 0 , lRdSolde , FInfo.Devise.Dev )
     else
      CSetMontants( lTOB , lRdSolde * (-1) , 0 , FInfo.Devise.Dev ) ;

   CVentilerTOB(lTOB,FInfo) ;

   if FInfo.Compte.IsLettrable then
    CRemplirInfoLettrage(lTOB)
     else
      CSupprimerInfoLettrage(lTOB) ;

   if FInfo.Compte.IsPointable then
    CRemplirInfoPointage(lTOB) ;

   CGetTVA(lTOB,FInfo) ;
   CGetEch(lTOB,FInfo) ;
  end ;

 CAffectRegimeTva(FTOBEcr) ;
 CAffectCompteContrePartie(FTOBEcr,FInfo) ;

 result := TRecError(CIsValidSaisiePiece ( FTOBEcr , FInfo )).RC_Error ;

end ;

procedure TOF_CPMULGCD.TransactionPieceFin ;
begin
 V_PGI.IoError := oeOk ;
 if not CEnregistreSaisie(FTOBEcr, True, False, True, FInfo ) then
  V_PGI.IoError := oeSaisie ;
end ;



procedure TOF_CPMULGCD.OnClickPOP8 ;
var
 lMessCompta : TMessageCompta ;        // affichage des messages
 lInError    : integer ;
// lTOBEcr     : TOB ;
// lQ          : TQuery ;
// lTOB        : TOB ;
begin

 if PGIAsk('Confirmez vous la génération des écritures ?') = MrNo then exit ;

 if ( GetParamSocSecur('SO_CPGCDREPRISE' , '' ) = '' ) then
  begin
   PGIInfo('Le compte de reprises n''est pas défini' ) ;
   exit ;
  end ;

 if ( GetParamSocSecur('SO_CPGCDDOTPROV' , '' ) = '' ) then
  begin
   PGIInfo('Le compte de dotation n''est pas défini' ) ;
   exit ;
  end ;

 if ( GetParamSocSecur('SO_CPGCDPROVISION' , '' ) = '' ) then
  begin
   PGIInfo('Le compte de provision n''est pas défini' ) ;
   exit ;
  end ;

 FTOBECR         := TOB.Create( '' , nil , -1 ) ;
 lMessCompta     := TMessageCompta.Create('Créance douteuses');

 try

  lInError := ContruirePieceFin ;
  if lInError = RC_BADWRITE then
   exit
    else
     if lInError = RC_PASERREUR then
      begin
       Transactions(TransactionPieceFin,1) ;
       _OuvrirPiece( FTOBECR ) ;
      end
       else
        lMessCompta.Execute(lInError) ;

 finally
  FreeAndNil(FTOBEcr) ;
  FreeAndNil(lMessCompta) ;
 end ;

end ;


procedure TOF_CPMULGCD.GCDCompta ;
var
 lQ             : TQuery ;
 lRdPerteHT     : double ;
 lRdSoldeDeb    : double ;
 lRdSoldeFin    : double ;
 lRdSolde654    : double ;
 lRdSolde491P   : double ;
 lRdSolde491E   : double ;
 lRdSolde681    : double ;
 lRdSolde781    : double ;
begin

 BAgrandir.Visible := false ;
 BReduire.Visible  := false ;
 PopF11.OnPopup    := nil ;
 SetControlVisible('BINSERT', false) ;
 SetControlVisible('BOUTIL', false) ;
 SetControlVisible('BTRAITEMENT', false) ;
 SetControlVisible('BPARAMLISTE', false) ;
 SetControlVisible('BIMPORT', false) ;
 SetControlVisible('BDELETE', false) ;
 SetControlVisible('BEXPORT', false) ;
 SetControlVisible('BSELECTALL', false) ;
// SetControlVisible('BSELECTALL', false) ;
 Ecran.Caption := TraduireMemoire('Gestion des créances douteuses ') ;
 UpdateCaption(Ecran);

 BAgrandir.Visible := False;
 BReduire.Visible  := True;
 BCherche.OnClick  := nil ;
 lRdSoldeDeb       := 0 ;
 lRdSoldeFin       := 0 ;
 lRdSolde654       := 0 ;
 lRdSolde491P      := 0 ;
 lRdSolde491E      := 0 ;
 lRdSolde681       := 0 ;
 lRdSolde781       := 0 ;
 lQ                := OpenSQL('select sum(GOP_HT) N from CPGCDOPERATION where GOP_CATEGORIE=4' , true) ;
 lRdPerteHT        := lQ.FindField('N').asFloat ;
 Ferme(lQ) ;

 if FInfo.LoadCompte(GetParamSocSecur('SO_CPGCDGENERAL', '' ) , true ) then
  begin
   lRdSoldeDeb  := Arrondi( FInfo.GetValue('G_TOTDEBP') - FInfo.GetValue('G_TOTCREP') , 2 ) ;
   lRdSoldeFin  := Arrondi( FInfo.GetValue('G_TOTDEBE') - FInfo.GetValue('G_TOTCREE') , 2 ) ;
  end ;

 if FInfo.LoadCompte(GetParamSocSecur('SO_CPGCDPERTE', '' ) ,true ) then
  lRdSolde654 := Arrondi( FInfo.GetValue('G_TOTDEBE') - FInfo.GetValue('G_TOTCREE'), 2 ) ;

 if FInfo.LoadCompte(GetParamSocSecur('SO_CPGCDPROVISION', '' ) , true) then
  begin
   lRdSolde491P := Arrondi( FInfo.GetValue('G_TOTDEBP') - FInfo.GetValue('G_TOTCREP') , 2 ) ;
   lRdSolde491E := Arrondi( FInfo.GetValue('G_TOTDEBE') - FInfo.GetValue('G_TOTCREE') , 2 ) ;
  end ;

 if FInfo.LoadCompte(GetParamSocSecur('SO_CPGCDDOTPROV', '' ) , true ) then
  lRdSolde681 := Arrondi( FInfo.GetValue('G_TOTDEBE') - FInfo.GetValue('G_TOTCREE') , 2 ) ;

 if FInfo.LoadCompte(GetParamSocSecur('SO_CPGCDREPRISE', '' ) , true ) then
  lRdSolde781 := Arrondi( FInfo.GetValue('G_TOTDEBE') - FInfo.GetValue('G_TOTCREE') , 2 ) ;

 HGBeginUpdate(FListe) ;

 try

  FListe.OnDblClick  := nil ;
  FListe.VidePile(false) ;
  FListe.RowCount    := 8 ;
  FListe.ColCount    := 5 ;
  Fliste.DBIndicator := false ;
  FStListeParam      := '';

  FListe.Cells[0,0]   := TraduireMemoire('Compte') ;
  FListe.Cells[1,0]   := TraduireMemoire('Libellé') ;
  FListe.Cells[2,0]   := TraduireMemoire('Solde GCD') ;
  FListe.Cells[3,0]   := TraduireMemoire('Solde Compta') ;
  FListe.Cells[4,0]   := TraduireMemoire('Ecart') ;

  FListe.ColFormats[0] := '' ;
  FListe.ColFormats[1] := '' ;
  FListe.ColFormats[2] := '' ;
  FListe.ColFormats[3] := '' ;
  FListe.ColFormats[4] := '' ;
  FListe.ColFormats[5] := '' ;
  FListe.ColFormats[6] := '' ;
  FListe.ColFormats[7] := '' ;

  FListe.Cells[0,1]   := GetParamSocSecur('SO_CPGCDGENERAL', '' ) ;
  FListe.Cells[0,2]   := GetParamSocSecur('SO_CPGCDGENERAL', '' ) ;
  FListe.Cells[0,3]   := GetParamSocSecur('SO_CPGCDPERTE', '' ) ;
  FListe.Cells[0,4]   := GetParamSocSecur('SO_CPGCDDOTPROV', '' ) ;
  FListe.Cells[0,5]   := GetParamSocSecur('SO_CPGCDREPRISE', '' ) ;
  FListe.Cells[0,6]   := GetParamSocSecur('SO_CPGCDPROVISION', '' ) ;
  FListe.Cells[0,7]   := GetParamSocSecur('SO_CPGCDPROVISION', '' ) ;

  FListe.Cells[1,1]   := TraduireMemoire('Début exercice') ;
  FListe.Cells[1,2]   := TraduireMemoire('Fin d''exercice' ) ;
  FListe.Cells[1,3]   := TraduireMemoire('') ;
  FListe.Cells[1,4]   := TraduireMemoire('') ;
  FListe.Cells[1,5]   := TraduireMemoire('') ;
  FListe.Cells[1,6]   := TraduireMemoire('Début d''exercice') ;
  FListe.Cells[1,7]   := TraduireMemoire('Fin d''excercice') ;

  FListe.Cells[2,1]   := AfficheDBCR( GetCumulChamps('CGC_SOLDE')       , '#,##0.00') ;
  FListe.Cells[2,2]   := AfficheDBCR( GetCumulChamps('FINEXO')          , '#,##0.00') ;
  FListe.Cells[2,3]   := AfficheDBCR( lRdPerteHT                        , '#,##0.00') ;
  FListe.Cells[2,4]   := AfficheDBCR( GetCumulChamps('DOTATIONS')       , '#,##0.00') ;
  FListe.Cells[2,5]   := AfficheDBCR( -GetCumulChamps('REPRISES')        , '#,##0.00') ;
  FListe.Cells[2,6]   := AfficheDBCR( -GetCumulChamps('CGC_PROVISION')   , '#,##0.00') ;
  FListe.Cells[2,7]   := AfficheDBCR( -GetCumulChamps('FINEXOP')         , '#,##0.00') ;

  FListe.Cells[3,1]   := AfficheDBCR( lRdSoldeDeb                       , '#,##0.00') ;
  FListe.Cells[3,2]   := AfficheDBCR( lRdSoldeFin                       , '#,##0.00') ;
  FListe.Cells[3,3]   := AfficheDBCR( lRdSolde654                       , '#,##0.00') ;
  FListe.Cells[3,4]   := AfficheDBCR( lRdSolde681                       , '#,##0.00') ;
  FListe.Cells[3,5]   := AfficheDBCR( lRdSolde781                       , '#,##0.00') ;
  FListe.Cells[3,6]   := AfficheDBCR( lRdSolde491P                      , '#,##0.00') ;
  FListe.Cells[3,7]   := AfficheDBCR( lRdSolde491E                      , '#,##0.00') ;

  FListe.Cells[4,1]   := AfficheDBCR( GetCumulChamps('CGC_SOLDE') - lRdSoldeDeb       , '#,##0.00') ;
  FListe.Cells[4,2]   := AfficheDBCR( GetCumulChamps('FINEXO')    - lRdSoldeFin       , '#,##0.00') ;
  FListe.Cells[4,3]   := AfficheDBCR( lRdPerteHT + lRdSolde654                        , '#,##0.00') ;
  FListe.Cells[4,4]   := AfficheDBCR( lRdSolde681 - GetCumulChamps('DOTATIONS')       , '#,##0.00') ;
  FListe.Cells[4,5]   := AfficheDBCR( lRdSolde781 + GetCumulChamps('REPRISES')        , '#,##0.00') ;
  FListe.Cells[4,6]   := AfficheDBCR( lRdSolde491P + GetCumulChamps('CGC_PROVISION')  , '#,##0.00') ;
  FListe.Cells[4,7]   := AfficheDBCR( lRdSolde491E + GetCumulChamps('FINEXOP')        , '#,##0.00') ;

  FListe.Refresh ;
  TFVierge(Ecran).HMTrad.ResizeGridColumns(FListe) ;

 finally
  HGEndUpdate(FListe) ;
  FRappro             := true ;
  FPanelCumul.Visible := false ;
 end ;

end ;


procedure TOF_CPMULGCD.OnClickBAgrandir(Sender: Tobject);
begin
 GCDCompta ;
end;

procedure TOF_CPMULGCD.OnRowEnterFListe(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var
 lTOB : TOB ;
begin

 inherited ;

 lTOB := GetO(FListe) ;
 if lTOB = nil then exit ;

 FInfo.LoadAux(lTOB.GetString('CGC_AUXILIAIRE')) ;
 Ecran.Caption := 'Gestion des créances douteuses : ' + FInfo.GetString('T_LIBELLE') ;
 UpdateCaption(Ecran);

end;

function TOF_CPMULGCD.AfterLoad  : boolean ;
var
 lCancel : boolean ;
begin
 result := inherited AfterLoad ;
 OnRowEnterFListe(self,FListe.Row,lCancel,lCancel) ;
end ;


procedure TOF_CPMULGCD.OnClickBReduire(Sender: Tobject);
begin
 BReduire.Visible    := False;
 BAgrandir.Visible   := True;
 FListe.Cells[0,0]   := '' ;
 FListe.VidePile(false) ;
 FStListeParam       := 'CPGCDCUMULS';
 FListe.DBIndicator  := true ;
 FListe.ColCount     := FColCount ;
 FListe.OnDblClick   := FListeOnDblClick ;
 FRappro             := false ;
 FPanelCumul.Visible := true ;
 TFVierge(Ecran).HMTrad.ResizeGridColumns(FListe) ;
 BCherche.OnClick    := OnClickBCherche ;
 PopF11.OnPopup      := OnPopUpPopF11 ;
 RefreshPclPge ;
 SetControlVisible('BINSERT'     , true ) ;
 SetControlVisible('BTRAITEMENT' , true ) ;
 SetControlVisible('BPARAMLISTE' , true ) ;
 SetControlVisible('BIMPORT'     , true ) ;
 SetControlVisible('BDELETE'     , true ) ;
 SetControlVisible('BEXPORT'     , true ) ;
 SetControlVisible('BSELECTALL'  , true ) ;
 InitPopUp ; // on affecte les menu. la fenetre a pu etre lancer directment en mode gcdcompta ( donc ss que l'on affecte els menu )
end;


procedure TOF_CPMULGCD.OnClickPOP9 ;
begin
 GCDCompta ;
end ;

procedure TOF_CPMULGCD.OnClickPOP10 ;
begin
 LanceEtat('E','CGC','CG1',true,false,false,nil,'','Gestion des créances douteuses',false) ;
end ;


{ TOF_CPGCDOPERATION }

procedure TOF_CPGCDOPERATION.OnArgument( S : String ) ;
var
 lStMode    : string ;
 lStModule  : string ;
 lStValeur  : string ;
begin

 inherited ;

 FMessCompta           := TMessageCompta.Create('Créance douteuses');

 lStModule             := ReadTokenSt(S) ;
 lStMode               := ReadTokenSt(S) ;
 lStValeur             := ReadTokenSt(S) ;

 FInfo                 := TInfoEcriture.Create() ;
 FGCD                  := TZGCD.Create(FInfo) ;
 FGCD.OnError          := OnErrorTOB ;
 FInfo.OnError         := OnErrorTOB ;
 
 if lStMode = 'PERTE' then
  FGCD.CATEGORIE := GCDPerte
   else
    if lStMode = 'PROVISION' then
     FGCD.CATEGORIE := GCDProv
      else
       FGCD.CATEGORIE := GCDNew ;

 if lStModule = 'CONSULT' then
  FGCD.LoadFromTOB(TOB(TheData))
   else
    if ( lStModule = 'TAB' ) then
     FGCD.Load(lStValeur)
      else
       if ( lStModule = 'AUX' ) then
        begin
         FGCD.AUXILIAIRECONTRE  := lStValeur ;
         FGCD.GENERALCONTRE     := ReadTokenSt(S) ;
         FGCD.Load(lStValeur) ;
         FGCD.TTC             := Valeur(ReadTokenSt(S)) ;
         FGCD.FaireEcr        := true ;
        end
         else
          if ( lStModule = 'GEN' ) then
           begin
            FGCD.GENERALCONTRE  := ReadTokenSt(S) ;
            FGCD.TTC            := Valeur(ReadTokenSt(S)) ;
            FGCD.FaireEcr       := true ;
           end ;

 InitControl ;

 SetControlProperty('BDELETE' , 'Visible' , false );

 THEdit(GetControl('GOP_TTC' , true)).OnExit                      := OnExitTTC ;
 THEdit(GetControl('GOP_TVA' , true)).OnExit                      := OnExitTVA ;
 THEdit(GetControl('GOP_DATECOMPTABLE' , true)).OnExit            := OnExitDate ;
 THEdit(GetControl('GOP_PROVISION' , true)).OnExit                := OnExitProvision ;
 THEdit(GetControl('GOP_TAUX' , true)).OnExit                     := OnExitTaux ;
 THValComboBox(GetControl('GOP_CODETAUX' , true)).OnChange        := OnChangeCodeTaux ;
 THEdit(GetControl('GOP_AUXILIAIRE' , true)).OnExit               := OnExitAux ;
 THEdit(GetControl('GOP_LIBELLE' , true)).OnExit                  := OnExitLibelle ;
 TToolBarButton97(GetControl('BPARAM' , true)).OnClick            := OnClickBPARAM ;
 FOldKeyDown                                                      := Ecran.OnKeyDown ;
 Ecran.OnKeyDown                                                  := OnKeyDownEcran ;

 SetFocusControl('GOP_DATECOMPTABLE') ;

 VH^.Lequel := 'T_NATUREAUXI="CLI"' ;

end;

procedure TOF_CPGCDOPERATION.OnClose;
begin

 VH^.Lequel := '' ;

 FreeAndNil(FGCD) ;
 FreeAndNil(FInfo) ;
 FreeAndNil(FMessCompta) ;

 inherited;

end;

procedure TOF_CPGCDOPERATION.OnKeyDownEcran( Sender : TObject; var Key : Word; Shift : TShiftState ) ;
begin

 if csDestroying in Ecran.ComponentState then Exit ;

 case Key of
 VK_F6     : begin
              Key := 0 ;
              CalculSolde ;
             end ;
 VK_ESCAPE : begin Key := 0 ; Close ; end ;
 else
  FOldKeyDown(Sender,Key,Shift) ;
 end ;
end ;

procedure TOF_CPGCDOPERATION.OnErrorTOB ( sender : TObject ; Error : TRecError) ;
begin

 case Error.RC_Error of
  RC_TMONODEVISE : SetFocusControl('GOP_TTC') ;
 end ;

 if ( trim(Error.RC_Message) <> '' ) then
  PGIInfo(Error.RC_Message + #10#13 + Error.RC_Methode , Ecran.Caption )
   else
    if Error.RC_Error <> RC_PASERREUR  then
      FMessCompta.Execute(Error.RC_Error) ;
end ;


procedure TOF_CPGCDOPERATION.OnExitDate( Sender : TObject ) ;
begin
 if csDestroying in Ecran.ComponentState then Exit ;
 FGCD.DATECOMPTABLE := StrToDate(GetControlText('GOP_DATECOMPTABLE')) ;
 InitControl ;
end ;


procedure TOF_CPGCDOPERATION.OnExitTTC ( Sender : TObject ) ;
var
 lRdVal : double ;
begin
 if csDestroying in Ecran.ComponentState then Exit ;
 lRdVal := Valeur( GetControlText('GOP_TTC') ) ;
 if (FGCD.CATEGORIE = GCDPerte ) or (FGCD.CATEGORIE = GCDRegl ) then
  FGCD.TTC := - lRdVal
  else
   FGCD.TTC := lRdVal ;
 InitControl ;
end ;


procedure TOF_CPGCDOPERATION.OnExitTVA ( Sender : TObject ) ;
begin
 if csDestroying in Ecran.ComponentState then Exit ;
 FGCD.TTC     := Valeur( GetControlText('GOP_TTC') ) ;
 FGCD.TVA     := Valeur( GetControlText('GOP_TVA') ) ;
 FGCD.HT      := Arrondi(FGCD.TTC / ( 1 + ( FGCD.TVA / 100 ) ),2) ;
 SetControlText('GOP_HT'  , AfficheDBCR(FGCD.HT) ) ;
end ;

procedure TOF_CPGCDOPERATION.OnExitAux ( Sender : TObject ) ;
begin
 if csDestroying in Ecran.ComponentState then Exit ;
 FGCD.Load(GetControlText('GOP_AUXILIAIRE')) ;
 InitControl ;
 if GetControlText('GOP_AUXILIAIRE') = '' then
  THEdit(GetControl('GOP_AUXILIAIRE' , true)).ElipsisClick(Sender) ;
end ;

procedure TOF_CPGCDOPERATION.OnExitLibelle ( Sender : TObject ) ;
begin
 if csDestroying in Ecran.ComponentState then Exit ;
 FGCD.LIBELLE := GetControlText('GOP_LIBELLE') ;
// SetControlText('GOP_LIBELLE'        , FGCD.LIBELLE ) ;
end ;


procedure TOF_CPGCDOPERATION.OnExitProvision ( Sender : TObject ) ;
begin

 if csDestroying in Ecran.ComponentState then Exit ;

 FGCD.PROVISION := Valeur( GetControlText('GOP_PROVISION') ) ;

 SetControlText('GOP_PROVISION'  , AfficheDBCR(FGCD.PROVISION) ) ;
 SetControlProperty('GOP_TAUX', 'Value' , FGCD.TAUX / 100 ) ;

end ;

procedure TOF_CPGCDOPERATION.OnExitTaux ( Sender : TObject ) ;
begin

 if csDestroying in Ecran.ComponentState then Exit ;

 FGCD.TAUX := Valeur( GetControlText('GOP_TAUX') ) ;

 if FGCD.TAUX <> FGCD.TAUXBASE then
  begin
   if GetControlText('GOP_CODETAUX') <> '' then
    SetControlText('GOP_CODETAUX' , '' ) ;
  end ;

 SetControlText('GOP_PROVISION'  , AfficheDBCR(FGCD.PROVISION) ) ;
 SetControlProperty('GOP_TAUX', 'Value' , FGCD.TAUX / 100 ) ;

end ;

procedure TOF_CPGCDOPERATION.OnChangeCodeTaux ( Sender : TObject ) ;
var
 lQ      : TQuery ;
 lStCode : string ;
begin

 if csDestroying in Ecran.ComponentState then Exit ;

 lStCode        := GetControlText('GOP_CODETAUX') ;

 if Trim(lStCode) = '' then exit ;

 FGCD.CODETAUX  := lStCode ;
 FGCD.TAUXBASE  := 0 ;

 lQ             := OpenSQL('select GCZ_TAUX from CPGCDTAUX where GCZ_CODE="' + FGCD.CODETAUX +'" ' , true ) ;
 FGCD.TAUXBASE  := lQ.FindField('GCZ_TAUX').asFloat ;
 Ferme(lQ) ;

 FGCD.TAUX      := FGCD.TAUXBASE ;
 FGCD.PROVISION := FGCD.HT * ( FGCD.TAUX / 100 ) ;

 SetControlText('GOP_PROVISION'  , AfficheDBCR(FGCD.PROVISION) ) ;
 SetControlProperty('GOP_TAUX', 'Value' , FGCD.TAUX / 100 ) ;

 InitControl ;

end ;

procedure TOF_CPGCDOPERATION.OnClickBPARAM ( Sender : TObject ) ;
begin
 AGLLanceFiche('CP', 'CPGCDTAUX','','', 'ACTION=CREATION' ) ;
 RemplirListe('CPGCDTAUX','');
 THValComboBox(GetControl('GOP_CODETAUX')).ReLoad ;
end ;

procedure TOF_CPGCDOPERATION.InitControl ;
var
 lStCaption : string ;
 lBoProv     : boolean ;
begin

// FGCD.CalculCategorie ;

 SetControlText('GOP_AUXILIAIRE'     , FGCD.AUXILIAIRE ) ;
 SetControlText('GOP_LIBELLE'        , FGCD.LIBELLE ) ;
 SetControlText('GOP_DATECOMPTABLE'  , DateToStr(FGCD.DATECOMPTABLE) ) ;
 SetControlText('GOP_TTC'            , AfficheDBCR(FGCD.TTC) ) ;
 SetControlText('GOP_HT'             , AfficheDBCR(FGCD.HT) ) ;
 SetControlProperty('GOP_TVA'        ,'Value' , FGCD.TVA / 100 ) ;

 case FGCD.Categorie of
  GCDDeb     : lStCaption := TraduireMemoire('Créance antérieure') ;
  GCDNew     : lStCaption := TraduireMemoire('Nouvelle créance sur l''exercice') ;
  GCDProv    : lStCaption := TraduireMemoire('Provision antérieure') ;
  GCDRegl    : lStCaption := TraduireMemoire('Règlement') ;
  GCDPerte   : lStCaption := TraduireMemoire('Perte') ;
 end ; // case

 SetControlText('GOP_CATEGORIE',lStCaption) ;

 //lBoProv := ( ( FGCD.Categorie = GCDDeb ) and ( FGCD.CODETAUX = '' ) ) or ( FGCD.Categorie = GCDProv ) ;
 lBoProv := ( FGCD.Categorie = GCDDeb ) or ( FGCD.Categorie = GCDProv ) ;

 SetControlEnabled( 'GOP_PROVISION'   , lBoProv ) ;
 SetControlEnabled( 'GOP_TAUX'        , lBoProv ) ;
 SetControlEnabled( 'TGOP_PROVISION'  , lBoProv ) ;
 SetControlEnabled( 'TGOP_TAUX'       , lBoProv ) ;
 SetControlEnabled( 'GOP_CODETAUX'    , lBoProv ) ;  // FGCD.Categorie = GCDDeb ) ;
 SetControlEnabled( 'TGOP_CODETAUX'   , lBoProv ) ;  //FGCD.Categorie = GCDDeb ) ;
 SetControlEnabled( 'BPARAM'          , lBoProv ) ;  //FGCD.Categorie = GCDDeb ) ;
 SetControlEnabled( 'GOP_TTC'         , FGCD.Categorie <> GCDProv ) ;
 //SetControlEnabled( 'GOP_HT'          , FGCD.Categorie <> GCDProv ) ;
 SetControlEnabled( 'GOP_TVA'         , FGCD.Categorie <> GCDProv ) ;

 if ( FGCD.Categorie <> GCDDeb ) and ( FGCD.Categorie <> GCDProv ) then
  begin
   SetControlText('GOP_PROVISION' , '' ) ;
   SetControlText('GOP_TAUX'      , '' ) ;
  end ; // if

end ;

procedure TOF_CPGCDOPERATION.CalculSolde ;
begin
// if FGCD.CATEGORIE = GCDRegl then
//  begin
   FGCD.TTC := FGCD.Solde ;
   initControl ;
//  end ;
end ;

procedure TOF_CPGCDOPERATION.TransactionCreance ;
begin

 V_PGI.IoError := oeOk ;

 if not FGCD.EnregistreCreance then
  V_PGI.IoError := oeSaisie ;

end ;

procedure TOF_CPGCDOPERATION.TransactionLettrage ;
begin

 V_PGI.IoError := oeOk ;

 try

 if not FGCD.Lettrage then
  V_PGI.IoError := oeSaisie ;

 except
  On E : Exception do
   PGIError('Erreur lors du lettrage : ' + E.Message ) ;
 end ;

end ;


procedure TOF_CPGCDOPERATION.TransactionPiece ;
begin

 V_PGI.IoError := oeOk ;

 if not FGCD.EnregistrePiece then
  V_PGI.IoError := oeSaisie ;

end ;


procedure TOF_CPGCDOPERATION.OnUpdate ;
var
 lInError : integer ;
begin

 Ecran.ModalResult := mrNone ;
 FGCD.LIBELLE      := GetControlText('GOP_LIBELLE') ;

 lInError          := FGCD.IsValid ;

 if lInError <> RC_PASERREUR then
  begin
   case lInError of
    RC_TMONODEVISE ,
    RC_MONTANTINEXISTANT     : SetFocusControl('GOP_TTC') ;
    RC_AUXINEXISTANT         : SetFocusControl('GOP_AUXILIAIRE') ;
    else
     SetFocusControl('GOP_DATECOMPTABLE') ;
   end ; // case
   exit ;
  end ;

 if FGCD.AUXILIAIREcontre = '' then
  FGCD.AUXILIAIREcontre := GetControlText('GOP_AUXILIAIRE') ;

 if not FGCD.ConstruirePiece then exit ;

 Transactions( TransactionCreance , 0 ) ;

 if ( V_PGI.IoError = oeOk ) and FGCD.AGenerer then
  Transactions( TransactionPiece , 0 ) ;

 if ( V_PGI.IoError = oeOk ) and FGCD.ALettrer then
  Transactions( TransactionLettrage , 0 ) ;

 if V_PGI.IoError = oeOk then
  begin
   if FGCD.AGenerer then
    _OuvrirPiece(FGCD.Ecr) ;
   Ecran.ModalResult := mrOk ;
  end ;


 FGCD.RAZCumul ;

end ;

{ TZGCD }

constructor TZGCD.Create( vInfoEcr : TInfoEcriture ) ;
begin

 _ControleParamSoc ;  

 inherited Create(vInfoEcr) ;

 FGOP                       := TOB.Create('', nil , -1 ) ;
 FTOBCumul                  := TOB.Create('CPGCDCUMULS', nil , -1 ) ;
 FBoALettrer                := false ;
 FBoAGenerer                := false ;
 FAUXILIAIRE                := '' ;
 FGENERAL                   := '' ;
 FJOURNAL                   := GetParamSocSecur('SO_CPGCDJAL' , '') ;
 FGENERALCREANCE            := GetParamSocSecur('SO_CPGCDGENERAL' , '') ;
 FNATUREPIECE               := 'OD';
 FLIBELLE                   := '' ;
 FLIBELLEECR                := '' ;
 FCategorie                 := GCDDeb ;
 FDATECOMPTABLE             := ctxExercice.EnCours.Deb ;
 FBoFaireEcr                := false ;
 FBoForceTaux               := false ;
 PTOBEcr                    := nil ;
 PTOBALettrer               := nil ;

end ;

destructor TZGCD.Destroy ;
begin
 FGOP.Free ;
 FTOBCumul.Free ;
end ;

procedure TZGCD.SetDate ( Value : TDateTime ) ;
var
 lExoDate : TExoDate ;
begin

 FDATECOMPTABLE := Value ;

 if IsValidDate <> RC_PASERREUR then
  begin
   FDATECOMPTABLE := ctxExercice.EnCours.Deb ;
   FCATEGORIE     := GCDNew ;
   exit ;
  end ;

 lExoDate := ctxExercice.QuelExoDate(Value) ;

 if lExoDate.Code = ctxExercice.Precedent.Code then
  begin
   if ( FCATEGORIE = GCDRegl ) or ( FCATEGORIE = GCDPerte ) then
    FDATECOMPTABLE := ctxExercice.EnCours.Deb
     else
      FCATEGORIE := GCDDeb ;
  end
   else
    if ( FCATEGORIE = GCDDeb ) then
     FCATEGORIE := GCDNew ;

end ;

procedure TZGCD.SetCategorie ( Value : TTypeCreance ) ;
begin
 FCategorie  := Value ;
 if FCategorie = GCDPerte then
  FBoAGenerer := true ;
end ;



procedure TZGCD.SetTTC ( Value : double ) ;
begin
 if ( CATEGORIE = GCDPerte ) and ( Value > 0 ) then
  Value := Value * (-1) ;
 FTTC := Value ;
 if IsValidTTC = RC_PASERREUR then
  FHT  := Arrondi(FTTC / ( 1 + ( FTVA / 100 ) ),2)
   else
    FHT  := 0 ;
end ;

procedure TZGCD.SetTaux ( Value : double ) ;
begin

 FTAUX        := Value ;
 FBoForceTaux := true ;

 if FTAUX > 100 then
  FTAUX := 100 ;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 02/05/2007
Modifié le ... :   /  /
Description .. : - 02/05/2007 - FB ? - le calcul du taux etait double si on
Suite ........ : saisissai un creance anteraieu en meme temps que la
Suite ........ : provision
Mots clefs ... :
*****************************************************************}
procedure TZGCD.SetProvision ( Value : double ) ;
var
 lRdHT : double ;
begin

 FPROVISION := Value ;

 if FPROVISION > ( CumulHT - FTOBCumul.GetValue('CGC_PROVISION') ) then
  FPROVISION := CumulHT - FTOBCumul.GetValue('CGC_PROVISION') ;

 lRdHT := HT ;
 if lRdHT = 0 then
  lRdHT := CumulHT ;

 if ( PROVISION <> 0 ) and ( (HT+CumulHT) <> 0 ) and not FBoForceTaux then
  FTAUX := Arrondi ( (  PROVISION / (lRdHT) ) * 100 , 4 ) ;

end ;


function TZGCD.IsValidTTC : integer ;
begin

 result := RC_PASERREUR ;

 if ( ( Categorie = GCDDeb ) or ( Categorie = GCDNew ) )and ( TTC < 0 ) then
  begin
   NotifyError(RC_TMONODEVISE , 'Vous ne pouvez pas saisir une créance antérieur avec un solde négatif' ) ;
   TTC := 0 ;
   result := RC_TMONODEVISE ;
   exit ;
  end ;


{ if ( Categorie = GCDNew ) and ( Abs(TTC) > CumulTTC ) then
  begin
   NotifyError(RC_TMONODEVISE , 'Vous ne pouvez pas saisir une créance supérieur au solde de début d''exercice' );
   TTC := 0 ;
   result := RC_TMONODEVISE ;
   exit ;
  end ;   }

 if ( ( Categorie = GCDPerte ) or ( Categorie = GCDRegl ) ) and ( Abs(TTC) > Solde ) then
  begin
   NotifyError(RC_TMONODEVISE , 'Vous ne pouvez pas saisir une perte ou un règlement supérieur au solde de fin d''exercice' );
   TTC := 0 ;
   result := RC_TMONODEVISE ;
   exit ;
  end ;


end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 16/07/2007
Modifié le ... : 16/07/2007
Description .. : - LG - 16/07/2007 - FB 21070 - le calcul ds la date n'etait
Suite ........ : pas correct pour els exercice decalé
Mots clefs ... :
*****************************************************************}
function TZGCD.IsValidDate : integer ;
begin

 result := CIsValidDateC(FDATECOMPTABLE,Info,true) ;

 if ( result <> RC_PASERREUR ) then // and ( ctxExercice.Precedent.Code <> ctxExercice.QuelExoDate(FDATECOMPTABLE).Code ) then
  begin
   if FDATECOMPTABLE >= PlusDate(ctxExercice.EnCours.Deb,-1,'A') then
    result := RC_PASERREUR ;
  end ;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 10/07/2007
Modifié le ... :   /  /
Description .. : - LG - 10/07/2007 - FB 20563 - corection du msg
Mots clefs ... :
*****************************************************************}
function TZGCD.IsValidAux : integer ;
begin
 if not info.LoadAux(AUXILIAIRE) then
  begin
   result :=  RC_AUXINEXISTANT ;
   NotifyError('Le code de rattachement n''est pas renseigné ou est inconnu ' , '' , '' ) ;
   exit ;
  end ;
 result := RC_PASERREUR ;
end ;

function TZGCD.IsValid : integer ;
begin

 result := IsValidAux ;
 if result <> RC_PASERREUR then exit ;

 result := IsValidTTC ;
 if result <> RC_PASERREUR then exit ;

 result := IsValidDate ;
 if result <> RC_PASERREUR then
  begin
   NotifyError(result,'') ;
   exit ;
  end ;

 if ( Arrondi(TTC+PROVISION,2) = 0 ) then
  begin
   NotifyError(RC_MONTANTINEXISTANT , 'Vous ne pouvez pas saisir une créance sans montant' );
   result := RC_MONTANTINEXISTANT ;
  end ;

end ;


procedure TZGCD.LoadFromTOB( Value : TOB ) ;
var
 i            : integer ;
 lTOBRef      : TOB ;
 lBoLettrable : boolean ;
begin

 PTOBEcr          := Value ;
 lTOBRef          := PTOBEcr.Detail[0] ;

 Load(lTOBRef.GetString('E_AUXILIAIRE')) ;

 FGENERAL          := lTOBRef.GetString('E_GENERAL') ;
 FNATUREPIECE      := lTOBRef.GetString('E_NATUREPIECE') ;
 FGENERALCONTRE    := FGENERAL ;
 FAUXILIAIRECONTRE := lTOBRef.GetString('E_AUXILIAIRE') ;
 lBoLettrable      := true ;

 FBoAGenerer    := ( FJOURNAL <> '' ) and ( FGENERAL <> FGENERALCREANCE ) ;
 FTTC           := 0 ;

 for i := 0 to PTOBEcr.Detail.Count - 1  do
  begin
   FLIBELLEECR := PTOBEcr.Detail[i].GetValue('E_LIBELLE')  ;
   FTTC        := Arrondi( FTTC + PTOBEcr.Detail[i].GetValue('E_DEBIT') - PTOBEcr.Detail[i].GetValue('E_CREDIT') , 2 ) ;
   if FDATECOMPTABLE < PTOBEcr.Detail[i].GetValue('E_DATECOMPTABLE') then
    FDATECOMPTABLE := PTOBEcr.Detail[i].GetValue('E_DATECOMPTABLE') ;
   // on regarde si une des ecritures du paquets peu etre lettrée ( evite de controle le tier, general )
   lBoLettrable := lBoLettrable and ( PTOBEcr.Detail[i].GetValue('E_ECRANOUVEAU') = 'N' ) and
                   ( ( PTOBEcr.Detail[i].GetValue('E_ETATLETTRAGE') = 'AL' ) or ( PTOBEcr.Detail[i].GetValue('E_ETATLETTRAGE') = 'PL' ) ) ;
  end ; // if

 HT  := Arrondi(TTC / ( 1 + ( TVA / 100 ) ),2) ;

 if TTC < 0 then
  CATEGORIE := GCDRegl ;

 if FBoAGenerer and lBoLettrable then
  begin
   if not CControlePresenceLettrage(PTOBEcr) then
    begin
     FBoALettrer := false ;
     PGIInfo('Impossible de lettrer, la totalité des mouvements n''est pas incluse dans la sélection !','Lettrage');
     exit ;
    end // if
     else
      FBoALettrer := true ;
  end ;

end;

function TZGCD.GetLibelle  : string ;
begin
 result := FLIBELLEECR ;
 if result = '' then
  result := FLIBELLE ;
end ;

procedure TZGCD.SetLibelle ( Value : string ) ;
begin
 FLIBELLE    := Value ;
 FLIBELLEECR := '' ;
end ;


procedure TZGCD.Load( Value : string ) ;
var
 lStCompteTVA : string ;
 lStRegimeTva : string ;
 lStCompte    : string ;
 lRdTva       : double ;
begin

 if ( Value = '' ) or ( Value = AUXILIAIRE ) then exit ;

 FAUXILIAIRE := Value ;

 if AUXILIAIRE <> '' then
  begin

   Info.Load( GENERAL , AUXILIAIRE , JOURNAL ) ;

   if Info.Aux.InIndex = - 1 then
    begin
     FAUXILIAIRE := '' ;
     exit ;
    end
     else
      begin
       FAUXILIAIRE  := Info.StAux ;
       lStRegimeTva := Info.getString('T_REGIMETVA') ;
      end ;


   FLIBELLE := Info.GetString('T_LIBELLE') ;

   if FLIBELLE = '' then
    FLIBELLE := Info.GetString('G_LIBELLE') ;

   if GENERAL <>  '' then
    lStCompte := GENERAL
     else
      lStCompte := AUXILIAIRE ;

      Info.Compte.RecupInfoTVA( lStCompte ,
                           TTC,
                           NATUREPIECE ,
                           Info.GetString('J_NATUREJAL'),
                           Info.GetString('T_NATUREAUXI'),
                           lStCompteTVA,
                           lRdTva,
                           lStRegimeTva,
                           false ,
                           false ,
                           true
                             );
   TVA        := lRdTva ;
   FCOMPTETVA := lStCompteTVA ;
   if TVA <> 0 then
    HT := Arrondi(TTC / ( 1 + ( TVA / 100 ) ),2) ;
  end ;

 RAZCumul ;

end ;

procedure TZGCD.RAZCumul ;
var
 lQ : TQuery ;
begin
 FTOBCumul.ClearDetail ;
 lQ := OpenSQL ('select * from CPGCDCUMULS where CGC_AUXILIAIRE="' + AUXILIAIRE + '" ' , true ) ;
 FTOBCumul.SelectDB('' , lQ ) ;
 Ferme(lQ) ;
end ;


function TZGCD.GetCumulTTC : double ;
begin
 result := FTOBCumul.GetValue('CGC_SOLDE') ;
end ;

function TZGCD.GetCumulHT : double ;
begin
 result := FTOBCumul.GetValue('CGC_HT') + HT ;
end ;

function TZGCD.GetSolde : double ;
begin
 result := FTOBCumul.GetValue('CGC_SOLDE') + FTOBCumul.GetValue('CGC_CREANCE') - FTOBCumul.GetValue('CGC_REGLEMENT') - FTOBCumul.GetValue('CGC_PERTE')  ;
end ;

procedure TZGCD.MajCumuls ;
var
 lStChamps  : string ;
 lBoExist   : boolean ;
 lTOBCumul  : TOB ;
 lStSQLTaux : string ;
 lStSQLProv : string ;
begin

  case CATEGORIE of
   GCDDeb   : lStChamps := 'CGC_SOLDE' ;
   GCDNew   : lStChamps := 'CGC_CREANCE' ;
   GCDRegl  : lStChamps := 'CGC_REGLEMENT' ;
   GCDPerte : lStChamps := 'CGC_PERTE' ;
   GCDProv  : lStChamps := 'CGC_PROVISION' ;
 end ;

 lBoExist := ExisteSQL('select CGC_AUXILIAIRE from CPGCDCUMULS where CGC_AUXILIAIRE="' + AUXILIAIRE + '" ' ) ;

 if not lBoExist then
  begin
   lTOBCumul := TOB.Create('CPGCDCUMULS', nil , -1 ) ;
   lTOBCumul.PutValue( 'CGC_AUXILIAIRE'  , AUXILIAIRE ) ;
   lTOBCumul.PutValue( 'CGC_LIBELLE'     , LIBELLE ) ;
   lTOBCumul.PutValue( lStChamps         , Abs(TTC) ) ;
   lTOBCumul.PutValue( 'CGC_HT'          , HT ) ;
   lTOBCumul.PutValue( 'CGC_CODETAUX'    , CODETAUX ) ;
   lTOBCumul.PutValue( 'CGC_TAUX'        , TAUX ) ;
   lTOBCumul.PutValue( 'CGC_PROVISION'   , PROVISION ) ;
   lTOBCumul.InsertDB(nil) ;
   lTOBCumul.Free ;
  end
   else
    begin
     lStSQLTaux := '' ;
     lStSQLProv := '' ;
     if TAUX <> 0 then
      lStSQLTaux :=  'CGC_TAUX=' + VariantToSql(TAUX) + ', ' ;
     if ( PROVISION <> 0 ) and ( CATEGORIE = GCDDeb ) then
      lStSQLProv := 'CGC_PROVISION=CGC_PROVISION + ' + VariantToSql(PROVISION) + ', ' ;

     if CATEGORIE = GCDProv then
      ExecuteSQL('update CPGCDCUMULS set ' +
      'CGC_PROVISION=CGC_PROVISION + ' + VariantToSql(PROVISION) + ', ' +
      lStSQLTaux +
      'CGC_CODETAUX="' + CODETAUX +'" ' +
      ' where CGC_AUXILIAIRE="' + AUXILIAIRE + '" ' )
       else
         ExecuteSQL('update CPGCDCUMULS set ' +
                lStChamps + '= ' + lStChamps + '+ ' + VariantToSql(Abs(TTC)) + ', ' +
                'CGC_HT=CGC_HT +' + VariantToSql(HT) + ', ' +
                lStSQLProv +
                lStSQLTaux +
                'CGC_CODETAUX="' + CODETAUX +'" ' +
                ' where CGC_AUXILIAIRE="' + AUXILIAIRE + '" ' ) ;
    end ;

end ;



{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 29/08/2007
Modifié le ... :   /  /    
Description .. : - FB 21271 - correction pour les auxi trop court
Mots clefs ... : 
*****************************************************************}
function TZGCD.EnregistreCreance : boolean ;
var
 lTOB     : TOB ;
 lTOBCNTV : TOB ;
begin

 result        := false ;
 lTOB          := TOB.Create('CPGCDOPERATION' , nil , -1 ) ;
 lTOBCNTV      := TOB.Create('CNTVDETAIL' , nil , -1 ) ;

 try
 try

 lTOB.PutValue('GOP_NUMERO'         , _ProchainNumero ) ;
 lTOB.PutValue('GOP_AUXILIAIRE'     , Copy(AUXILIAIRE,1,6) ) ;
 lTOB.PutValue('GOP_DATECOMPTABLE'  , DATECOMPTABLE ) ;
 lTOB.PutValue('GOP_LIBELLE'        , LIBELLE ) ;
 lTOB.PutValue('GOP_TTC'            , TTC ) ;
 lTOB.PutValue('GOP_HT'             , HT ) ;
 lTOB.PutValue('GOP_TVA'            , TVA ) ;
 lTOB.PutValue('GOP_CATEGORIE'      , CATEGORIE ) ;
 lTOB.PutValue('GOP_CODETAUX'       , CODETAUX ) ;

 lTOB.InsertDB(nil) ;

 lTOBCNTV.PutValue('CND_NUMEROSEQ'   , lTOB.GetValue('GOP_NUMERO') ) ;
 lTOBCNTV.PutValue('CND_AUXILIAIRE'  , AUXILIAIRE ) ;
 lTOBCNTV.PutValue('CND_PERM'        , '1' ) ;
 lTOBCNTV.PutValue('CND_EXERCICE'    , ctxexercice.EnCours.Code ) ;
 lTOBCNTV.InsertDB(nil) ;

 if PROVISION <> 0 then
  begin
   lTOB.PutValue('GOP_NUMERO'        , _ProchainNumero ) ;
   lTOB.PutValue('GOP_TTC'           , PROVISION ) ;
   lTOB.PutValue('GOP_HT'            , PROVISION ) ;
   lTOB.PutValue('GOP_TVA'           , TAUX ) ;
   lTOB.PutValue('GOP_CATEGORIE'     , Integer(GCDProv) ) ;
   lTOB.InsertDB(nil) ;
   lTOBCNTV.PutValue('CND_NUMEROSEQ'  , lTOB.GetValue('GOP_NUMERO') ) ;
   lTOBCNTV.InsertDB(nil)
  end ;

 MajCumuls ;

 result := true ;

 except
  on E: Exception do
   begin
    PGIInfo('Erreur lors de l''enregistrement de la créance' + #10#13 + E.Message ) ;
   end ;
 end ;

 finally
  lTOB.Free ;
  lTOBCNTV.Free ;
 end ;

end ;

function TZGCD.EnregistrePiece : boolean ;
begin
 result := true ;
 if FBoAGenerer and ( FGOP.Detail.Count > 0 ) then
  result := CEnregistreSaisie(FGOP, True, False, True, Info ) ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 17/09/2007
Modifié le ... :   /  /    
Description .. : - LG - 17/09/2007 - FB 21007 - chg du libelle de generation
Mots clefs ... : 
*****************************************************************}
procedure TZGCD.ConstruirePieceLigneUn ;
var
 lTOB      : TOB ;
 lDtDate   : TDateTime ;
begin

 lDtDate := FDATECOMPTABLE ;
 if lDtDate < ctxExercice.EnCours.Deb then
  lDtDate := ctxExercice.EnCours.Deb ;

 Info.LoadJournal(FJOURNAL) ;
 Info.LoadCompte(FGENERALCREANCE) ;

 FGOP.ClearDetail ;

 lTOB    := TOB.Create('ECRITURE' , FGOP , - 1 ) ;
 CPutDefautEcr(lTOB) ;
 lTOB.PutValue('E_GENERAL'        , FGENERALCREANCE) ;
 if ( Info.GetValue('G_COLLECTIF') = 'X' ) then
  lTOB.PutValue('E_AUXILIAIRE'    , FAUXILIAIREcontre) ;
 lTOB.PutValue('E_NUMLIGNE'       , -1 ) ;
 lTOB.PutValue('E_LIBELLE'        , 'INSCRIPTION EN CREANCES DOUTEUSES' ) ; //FLIBELLE ) ;
 lTOB.PutValue('E_NUMEROPIECE'    , -1 ) ;
 lTOB.PutValue('E_NUMGROUPEECR'   , 0 ) ;
 lTOB.PutValue('E_JOURNAL'        , FJOURNAL ) ;
 lTOB.PutValue('E_MODESAISIE'     , Info.GetString('J_MODESAISIE') ) ;
 lTOB.PutValue('E_IO'             , 'X') ;
 lTOB.PutValue('E_QUALIFORIGINE'  , 'GCD' ) ;
 CRemplirDateComptable(lTOB , lDtDate ) ;

 if FTTC > 0 then
  CSetMontants( lTOB , FTTC, 0 , Info.Devise.Dev )
   else
    CSetMontants( lTOB , 0 , FTTC * (-1) , Info.Devise.Dev ) ;

 CVentilerTOB(lTOB,Info) ;

 if Info.Compte.IsLettrable then
  CRemplirInfoLettrage(lTOB)
   else
    CSupprimerInfoLettrage(lTOB) ;

 if Info.Compte.IsPointable then
  CRemplirInfoPointage(lTOB) ;

 CGetTVA(lTOB,Info ) ;
 CGetEch(lTOB,Info) ;

end ;


procedure TZGCD.ConstruirePiecePerte ;
var
 lTOB      : TOB ;
 lDtDate   : TDateTime ;
 lRdTVA    : double ;
begin

 lRdTVA  := Abs(TTC - HT) ;

 lDtDate := FDATECOMPTABLE ;
 if lDtDate < ctxExercice.EnCours.Deb then
  lDtDate := ctxExercice.EnCours.Deb ;

 ConstruirePieceLigneUn ;

 lTOB := TOB.Create('ECRITURE' , FGOP , - 1 ) ;
 CPutDefautEcr(lTOB) ;
 lTOB.PutValue('E_NUMLIGNE'       , -1 ) ;
 lTOB.PutValue('E_NUMEROPIECE'    , -1 ) ;
 lTOB.PutValue('E_NUMGROUPEECR'   , 0 ) ;
 lTOB.PutValue('E_LIBELLE'        , FLIBELLE ) ;
 lTOB.PutValue('E_GENERAL'        , GetParamSocSecur('SO_CPGCDPERTE', '') ) ;
 lTOB.PutValue('E_JOURNAL'        , FJOURNAL ) ;
 lTOB.PutValue('E_MODESAISIE'     , Info.GetString('J_MODESAISIE') ) ;
 lTOB.PutValue('E_IO'             , 'X') ;
 lTOB.PutValue('E_QUALIFORIGINE'  , 'GCD' ) ;
 CRemplirDateComptable(lTOB , lDtDate ) ;

 CSetMontants( lTOB , Abs(HT) , 0 , Info.Devise.Dev ) ;
 CVentilerTOB(lTOB,Info) ;

 if Info.Compte.IsLettrable then
  CRemplirInfoLettrage(lTOB)
   else
    CSupprimerInfoLettrage(lTOB) ;

 if Info.Compte.IsPointable then
  CRemplirInfoPointage(lTOB) ;

 CGetTVA(lTOB,Info) ;
 CGetEch(lTOB,Info) ;

 if FCOMPTETVA <> '' then
  begin
   lTOB := TOB.Create('ECRITURE' , FGOP , - 1 ) ;
   CPutDefautEcr(lTOB) ;
   lTOB.PutValue('E_NUMLIGNE'       , -1 ) ;
   lTOB.PutValue('E_NUMEROPIECE'    , -1 ) ;
   lTOB.PutValue('E_NUMGROUPEECR'   , 0 ) ;
   lTOB.PutValue('E_LIBELLE'        , FLIBELLE ) ;
   lTOB.PutValue('E_GENERAL'        , FCOMPTETVA ) ;
   lTOB.PutValue('E_JOURNAL'        , FJOURNAL ) ;
   lTOB.PutValue('E_MODESAISIE'     , Info.GetString('J_MODESAISIE') ) ;
   lTOB.PutValue('E_IO'             , 'X') ;
   lTOB.PutValue('E_QUALIFORIGINE'  , 'GCD' ) ;
   CRemplirDateComptable(lTOB , lDtDate ) ;

   CSetMontants( lTOB , lRdTVA , 0 , Info.Devise.Dev ) ;
   CVentilerTOB(lTOB,Info) ;

   if Info.Compte.IsLettrable then
    CRemplirInfoLettrage(lTOB)
     else
      CSupprimerInfoLettrage(lTOB) ;

   if Info.Compte.IsPointable then
    CRemplirInfoPointage(lTOB) ;

   CGetTVA(lTOB,Info) ;
   CGetEch(lTOB,Info) ;
  end ; // if

end ;

procedure TZGCD.ConstruirePieceCreance ;
var
 lTOB      : TOB ;
 lDtDate   : TDateTime ;
begin

 lDtDate := FDATECOMPTABLE ;
 if lDtDate < ctxExercice.EnCours.Deb then
  lDtDate := ctxExercice.EnCours.Deb ;

 ConstruirePieceLigneUn ;

 Info.LoadCompte(FGENERALCONTRE) ;

 lTOB := TOB.Create('ECRITURE' , FGOP , - 1 ) ;
 CPutDefautEcr(lTOB) ;
 lTOB.PutValue('E_NUMLIGNE'       , -1 ) ;
 lTOB.PutValue('E_NUMEROPIECE'    , -1 ) ;
 lTOB.PutValue('E_NUMGROUPEECR'   , 0 ) ;
 lTOB.PutValue('E_LIBELLE'        , FLIBELLE ) ;
 lTOB.PutValue('E_GENERAL'        , FGENERALCONTRE ) ;
 if ( Info.GetValue('G_COLLECTIF') = 'X' ) then
  lTOB.PutValue('E_AUXILIAIRE'     , FAUXILIAIRECONTRE ) ;
 lTOB.PutValue('E_JOURNAL'        , FJOURNAL ) ;
 lTOB.PutValue('E_MODESAISIE'     , Info.GetString('J_MODESAISIE') ) ;
 lTOB.PutValue('E_IO'             , 'X') ;
 lTOB.PutValue('E_QUALIFORIGINE'  , 'GCD' ) ;
 //CDupliquerInfoAux(lTOBRef,lTOB) ;
 CRemplirDateComptable(lTOB , lDtDate ) ;
 PTOBALettrer := lTOB ;

 if FTTC > 0 then
  CSetMontants( lTOB , 0 , FTTC , Info.Devise.Dev )
   else
    CSetMontants( lTOB , FTTC * (-1), 0 , Info.Devise.Dev ) ;

 CVentilerTOB(lTOB,Info) ;

 if Info.Compte.IsLettrable then
  CRemplirInfoLettrage(lTOB)
   else
    CSupprimerInfoLettrage(lTOB) ;

 if Info.Compte.IsPointable then
  CRemplirInfoPointage(lTOB) ;

 CGetTVA(lTOB,Info ) ;
 CGetEch(lTOB,Info) ;


end ;

function TZGCD.ConstruirePiece : boolean ;
var
 lRecError : TRecError ;
begin

 result := false ;

 if FBoFaireEcr and not FBoAGenerer then
  begin
   // on a demander a generer des ecritures de contrepartie -> test si on peut vraiment le faire
   FBoAGenerer    := ( FGENERALCONTRE <> FGENERALCREANCE ) ;
   if FBoAGenerer then
    begin
     Info.Load(FGENERALCONTRE,FAUXILIAIRECONTRE,'') ;
     FBoALettrer := Info.Compte.IsLettrable() or ( Info.GetString('T_LETTRABLE') = 'X' ) ;
    end ;
  end ;


 if not FBoAGenerer then
  result := true
   else
    begin

     Info.Devise.Load([V_PGI.DevisePivot]) ;

     if CATEGORIE = GCDPerte then
      ConstruirePiecePerte
       else
        ConstruirePieceCreance ;

     CAffectRegimeTva(FGOP) ;
     CAffectCompteContrePartie(FGOP,Info) ;

     lRecError := CIsValidSaisiePiece ( FGOP , Info ) ;
     if ( lRecError.RC_Error <> RC_PASERREUR ) then
      OnError(nil,lRecError)
        else
         result := true ;

    end ; // else
end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 25/07/2007
Modifié le ... : 03/10/2007
Description .. : - LG - FB 20970 - il mq les champs du lettrage
Suite ........ : - LG - FB 21521 - plantait depuis la consultations des 
Suite ........ : generaux
Mots clefs ... : 
*****************************************************************}
function TZGCD.Lettrage : boolean ;
var
 lOldParent : TOB ;
 lTOB       : TOB ;
begin

 result := true ;

 if not FBoALettrer or ( PTOBALettrer =  nil ) then exit ;
 if ( PTOBEcr = nil ) or ( PTOBEcr.Detail = nil ) then exit ;


 lOldParent := PTOBALettrer.Parent ;

 try

  PTOBALettrer.ChangeParent(PTOBEcr,-1) ;
  lTOB := PTOBEcr.Detail[0] ;
  lTOB.AddChampSupValeur ('OLDDEBIT'     , 0 , true );
  lTOB.AddChampSupValeur ('OLDCREDIT'    , 0 , true );
  lTOB.AddChampSupValeur ('RATIO'        , 0 , true );
  lTOB.AddChampSupValeur ('CONVERTFRANC' , 0 , true);
  lTOB.AddChampSupValeur ('CONVERTEURO'  , 0 , true );
  CLettrageTOB ( PTOBEcr ) ;
  result := true ;

 finally
  PTOBALettrer.ChangeParent(lOldParent,-1) ;
 end ;

end ;


procedure TOF_CPMULGCD.InitControl;
begin
end;

{ TOM_CPGCDTAUX }

procedure TOM_CPGCDTAUX.OnArgument( S : String);
var
 lEdit : THEdit ;
begin

 lEdit         := THedit(GetControl('GCZ_CODE',true) );
 lEdit.Libelle := nil ;
 lEdit         := THedit(GetControl('GCZ_DATEDU',true) );
 lEdit.Libelle := nil ;
 lEdit         := THedit(GetControl('GCZ_DATEAU',true) );
 lEdit.Libelle := nil ;
 lEdit         := THedit(GetControl('GCZ_LIBELLE',true) );
 lEdit.Libelle := nil ;

 inherited ;

 TToolBarButton97(GetControl('BVALIDER' , true)).OnClick := OnClickBValider ;

end;

procedure TOM_CPGCDTAUX.OnClickBValider ( Sender : TObject ) ;
begin
 TFFicheListe(Ecran).BValiderClick(Sender) ;
 TFFicheListe(Ecran).ModalResult := mrOk ;
 TFFicheListe(Ecran).Retour := GetField('GCZ_CODE') + ';' + VariantToSQL (GetField('GCZ_TAUX')) ;
end ;

procedure TOM_CPGCDTAUX.OnClose;
begin
 inherited ;
end;


procedure TOM_CPGCDTAUX.OnChangeField(F: TField);
begin
 inherited;
 if F.FieldName = 'GCZ_TAUX' then
  begin
    if GetField('GCZ_TAUX') > 100 then
     Setfield('GCZ_TAUX' , 100 ) ;
  end ;
end;


procedure TOM_CPGCDTAUX.OnDeleteRecord;
begin
  inherited;
end;

procedure TOM_CPGCDTAUX.OnLoadRecord;
begin
  inherited;
end;

procedure TOM_CPGCDTAUX.OnNewRecord ;
begin

 inherited ;

 SetField('GCZ_DATEDU', ctxExercice.EnCours.Deb) ;
 SetField('GCZ_DATEAU', ctxExercice.EnCours.Fin) ;

end;

procedure TOM_CPGCDTAUX.OnUpdateRecord;
begin
  inherited;
end;

{ TOF_CPGCDLISTEOPE }

function _decodecategorie ( Value : integer ) : string ;
begin
 case Value of
  Integer(GCDDeb)     : result := TraduireMemoire('Solde') ;
  Integer(GCDNew)     : result := TraduireMemoire('Nvll.') ;
  Integer(GCDProv)    : result := TraduireMemoire('Prov.') ;
  Integer(GCDRegl)    : result := TraduireMemoire('Règl.') ;
  Integer(GCDPerte)   : result := TraduireMemoire('Perte') ;
 end ; // case
end ;

procedure TOF_CPGCDLISTEOPE.OnArgument( S : String) ;
begin
 inherited;
 FAux := ReadTokenSt(S) ;
 FTOB := TOB.Create('',nil,-1) ;
 RemplirFListe() ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 29/08/2007
Modifié le ... :   /  /    
Description .. : - FB 21271 - correction pour les auxi trop court
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPGCDLISTEOPE.RemplirFListe ;
var
 lQ      : TQuery ;
 lTOB    : TOB ;
 i       : integer ;
begin

 if FAux = '' then exit ;

 FListe  := THGrid(GetControl('FLISTE', True)) ;
 lQ      := OpenSQL('select CPGCDOPERATION.*, GCZ_LIBELLE from CNTVDETAIL , CPGCDOPERATION left outer join CPGCDTAUX ' +
                    'on ( GOP_CODETAUX=GCZ_CODE ) ' +
                    'where CND_NUMEROSEQ=GOP_NUMERO ' +
                    ' AND CND_PERM="1" ' +
                    ' AND CND_AUXILIAIRE="' + FAux + '" '
                   , true ) ;

 FTOB.ClearDetail ;

 FTOB.LoadDetailDB('CPGCDOPERATION','','',lQ,true) ;
 Ferme(lQ) ;

 FListe.RowCount := FTOB.Detail.Count + 1 ;

 for i := 0 to FTOB.Detail.Count - 1 do
  begin
   lTOB                  := FTOB.Detail[i] ;
   FListe.Cells[0,i+1]   := lTOB.GetString('gop_datecomptable') ;
   FListe.Cells[1,i+1]   := _DecodeCategorie( lTOB.GetValue('gop_categorie') ) ;
   FListe.Cells[2,i+1]   := lTOB.GetString('gop_libelle') ;
   FListe.Cells[3,i+1]   := AfficheDBCR(lTOB.GetValue('gop_ht')) ;
   if lTOB.GetValue('gop_categorie') = Integer(GCDProv) then
    FListe.Cells[6,i+1]   := FormatFloat('0.00 %' , lTOB.Getvalue('gop_tva') )
     else
      FListe.Cells[4,i+1]   := FormatFloat('0.00 %' , lTOB.Getvalue('gop_tva') ) ;
   if lTOB.GetValue('gop_categorie') <> Integer(GCDProv) then
    FListe.Cells[5,i+1]   := AfficheDBCR(lTOB.GetValue('gop_ttc')) ;
   FListe.Cells[7,i+1]   := lTOB.GetValue('GCZ_LIBELLE') ;
   FListe.Objects[0,i+1] := lTOB ;
  end ; // for

 TFVierge(Ecran).HMTrad.ResizeGridColumns(FListe) ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 22/08/2007
Modifié le ... : 29/08/2007
Description .. : - LG - FB 21269 - 22/08/2007- correction de la maj des
Suite ........ : reglemetns
Suite ........ : - FB 21271 - correction pour les auxi trop court
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPGCDLISTEOPE.MajCumuls ;
var
 lInCat    : integer ;
 lStChamps : string ;
 lRdValue  : double ;
 lTOB      : TOB ;
begin

 lTOB := TOB(FListe.Objects[0,FRow]) ;

 if lTOB = nil then exit ;

 try

 ExecuteSQL('DELETE CPGCDOPERATION WHERE GOP_NUMERO=' + VariantToSql(lTOB.GetValue('GOP_NUMERO')) ) ;
 ExecuteSQL('DELETE CNTVDETAIL WHERE CND_PERM="1" and CND_NUMEROSEQ=' + VariantToSql(lTOB.GetValue('GOP_NUMERO')) ) ;


 lInCat := lTOB.GetValue('GOP_CATEGORIE') ;

 case lInCat of
   Integer(GCDDeb)   : begin lStChamps := 'CGC_SOLDE'       ; lRdValue := -lTOB.GetValue('GOP_TTC') ; end ;
   Integer(GCDNew)   : begin lStChamps := 'CGC_CREANCE'     ; lRdValue := -lTOB.GetValue('GOP_TTC') ; end ;
   Integer(GCDRegl)  : begin lStChamps := 'CGC_REGLEMENT'   ; lRdValue := +lTOB.GetValue('GOP_TTC') ; end ;
   Integer(GCDPerte) : begin lStChamps := 'CGC_PERTE'       ; lRdValue := lTOB.GetValue('GOP_TTC') ; end ;
   Integer(GCDProv)  : begin lStChamps := 'CGC_PROVISION'   ; lRdValue := -lTOB.GetValue('GOP_HT') ; end ;
   else
    lRdValue := 0 ;
 end ;

 if lInCat = Integer(GCDProv) then
  ExecuteSQL('update CPGCDCUMULS set ' +
            lStChamps + '= ' + lStChamps + '+ ' + VariantToSql(lRdValue) +
            ' where CGC_AUXILIAIRE="' + FAux + '" ' )
  else
   ExecuteSQL('update CPGCDCUMULS set ' +
               lStChamps + '= ' + lStChamps + '+ ' + VariantToSql(lRdValue) +
              ',CGC_HT=CGC_HT -' + VariantToSql(lTOB.GetValue('GOP_HT')) +
              ' where CGC_AUXILIAIRE="' + FAux + '" ' ) ;

 except
  On E : Exception do
   begin
    PGIInfo('Erreur de suppression ' + #10#13 + E.Message ) ;
   end ;
 end ;
end ;


procedure TOF_CPGCDLISTEOPE.OnDelete ;
var
 i : integer ;
begin

 if PGIAsk('Voulez-vous supprimer les créances selectionnées ?') = mrNo then exit ;

 if (FListe.nbSelected > 0) or (FListe.AllSelected) then
  begin
   for i := 1 to FListe.RowCount - 1 do
     if FListe.IsSelected(i) then
      begin
       FRow := i ;
       Transactions( MajCumuls, 1 ) ;
      end
  end
   else
    begin
     FRow := FListe.Row ;
     Transactions( MajCumuls , 1 ) ;
    end ;

  inherited ;

 RemplirFListe ;

end ;

procedure TOF_CPGCDLISTEOPE.OnClose;
begin
 FTOB.Free ;
 inherited;
end;

Initialization
 registerclasses ( [ TOF_CPMULGCD , TOF_CPGCDOPERATION , TOM_CPGCDTAUX , TOF_CPGCDLISTEOPE] ) ;


end.


