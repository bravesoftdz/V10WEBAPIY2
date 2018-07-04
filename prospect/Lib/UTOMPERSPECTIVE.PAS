unit Utomperspective;

interface

uses     StdCtrls,Controls,Classes,forms,sysutils,
         HCtrls,HEnt1,HMsgBox,UTOM,UTob,graphics,Grids,ParamSoc,FactUtil,
         Windows,M3FP,
         menus,HTB97,AglIsoflex, utilGC,
{$IFDEF EAGLCLIENT}
      MaineAGL,eFiche,UtileAGL,Spin
{$ELSE}
      db,{$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}DBCtrls,
      Fe_Main,Fiche,HDB
{$ENDIF}
			,uEntCommun
      ;

Const
   Encours : string = 'ENC';
   Perdue : string = 'PER';
   Signee : string = 'SIG';
   Abandonnee : string = 'ABA';

Type
TOM_PERSPECTIVES = Class (TOM)
    procedure OnNewRecord ; override ;
    procedure OnUpdateRecord ; override ;
    procedure OnAfterUpdateRecord ; override ;
    procedure OnLoadRecord ; override ;
    procedure OnLoadAlerte; override ;
    procedure OnChangeField (F : TField) ; override ;
    procedure OnClose ; override ;
    procedure OnArgument (Arguments : String ); override ;
    function GetInfoGed : string ;
private
  	PieceBTP,PieceNEG : string;
    { mng_fq012;10859 }
    AsInsert : boolean;
    ListeFinConc,ListeInitConc : string;
    StTiersConc : string;
    LesColonnes: string ;
    LesColonnesPiece: string ;
    LesColonnesAffaire: string ;
    GS : THGRID ;
    GP : THGRID ;
    GA : THGRID ;
    MontantProposition : double;
    Action   : TActionFiche ;
    TOBPersp,TobPieceOrigine,TobAffaire : TOB;
    NouvProp,Bprincipale,BVariante,BTermine,NoChangeProspect,NoChangeContact,ProchaineAction,SaisieProposition : Boolean;
    NoZoomDevis : Boolean; // pas de zoom devis si l'on vient d'un devis
    PerspPrincipale,iNumAction,InumeroContact : Integer;
    OldPrincipale: integer; { mng_090608_fq01210685 }
    stTiers,stProjet,stIntervenant,stArgForNextAction : string;
    stAffaire : string; // pour le cas où l'on vient de la fiche affaire, report du code sur la fiche action
    DerniereCreate   : integer;
    Bouge : Boolean; // ne pas créer d'historique sur l'instruction bouge
    // mng 02-03-01procedure MajIntervintPersp;
    ModifLot,bDroitModifTiers : boolean;
    StSQL : string;
    procedure NouvellePiece;
    procedure NouvelleAffaire;    
    procedure PerspGenereAction;
    procedure PiecePerspective(TypeAppel:String);
    procedure VariantePerspective (TypeAppel:String);
//    procedure TOBCopieChamp(FromTOB, ToTOB : TOB);
    procedure EnvoiMessage;
    procedure ChargeGrid( ForceCalcul:Boolean);
    procedure ChargeGridGA;    
    procedure ClickPrincipal;
    procedure GSLigneDClick (Sender: TObject);
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSCellExit(Sender: TObject; var ACol,
                                        ARow: Integer; var Cancel: Boolean);
    procedure GSCellEnter(Sender: TObject; var ACol,
                                        ARow: Integer; var Cancel: Boolean);
    procedure GSElipsisClick(Sender: TObject);
    procedure GSExit(Sender: TObject);
    procedure GSEnter(Sender: TObject);

    procedure DessineCellGS ( ACol,ARow : Longint; Canvas : TCanvas;
                                     AState: TGridDrawState);
    procedure DessineCellGP ( ACol,ARow : Longint; Canvas : TCanvas;
                                     AState: TGridDrawState);
    procedure GPLigneDClick (Sender: TObject);
    procedure GPRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GPRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
{$IFDEF AFFAIRE}
    procedure DessineCellGA ( ACol,ARow : Longint; Canvas : TCanvas;
                                     AState: TGridDrawState);
    procedure GALigneDClick (Sender: TObject);
    procedure GARowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GARowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
{$ENDIF}
    function AfficheTiers(ARow: Longint) : boolean;
    procedure ContactNom;
    function TestConcurrent ( ACol,ARow : integer; Elipsis : boolean) : boolean;
    // Gestion Isoflex
    procedure GereIsoflex;
    procedure AddRemoveItemFromPopup(stPopup,stItem : string; bVisible : boolean);
    procedure RTAppelGED_OnClick (Sender: TObject) ;
    procedure GestionBoutonGED;
    procedure SetArguments(StSQL : string);
    procedure ListAlerte_OnClick_RPE(Sender: TObject);
    procedure Alerte_OnClick_RPE(Sender: TObject);
    procedure RTPropoAppelParamCL (bCreation : boolean);
		procedure BIMPRIMEFICHE_OnClick(Sender:TObject);
    function GetTypePiece: string;
Public
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
END ;

procedure DessineCell (GG : Thgrid; ACol, ARow: Longint; Canvas : TCanvas; AState: TGridDrawState);

procedure AGLEnvoiMessage(parms:array of variant; nb: integer ) ;
procedure AGLChargeGrid(parms:array of variant; nb: integer ) ;
procedure AGLClickPrincipal (parms:array of variant; nb: integer ) ;
function Perspective_MyAfterImport (Sender : TObject): string ;
procedure Perspective_GestionBoutonGed (Sender : TObject);

const
	// libellés des messages
	TexteMessage: array[1..26] of string 	= (
          {1}        'Vous devez renseigner un code Tiers'
          {2}        ,'Le code client ou prospect n''existe pas'
          {3}        ,'La nature de la proposition doit être renseignée'
          {4}        ,'Le libellé de la proposition doit être renseigné'
          {5}        ,'L''opération d''origine n''existe pas'
          {6}        ,'La situation de la proposition doit être renseignée'
          {7}        ,'Vous devez renseigner le motif'
          {8}        ,'Suppression devis non autorisée'
          {9}        ,'Ce concurrent est déjà sélectionné'
          {10}       ,'Ce projet n''est pas multi-tiers'
          {11}       ,'Le tiers du projet associé doit être renseigné'
          {12}       ,'Le tiers du projet associé doit être du même groupe'
          {13}       ,'La proposition doit être rattachée à un projet'
          {14}       ,'Vous n''êtes pas autorisé à saisir cette proposition'
          {15}       ,'Le code du responsable doit être renseigné'
          {16}       ,'Le code du commercial doit être renseigné'
          {17}       ,'Le code responsable n''existe pas'
          {18}       ,'Le code commercial n''existe pas'
          {19}       ,'L''estimation réussite doit être au moins égale à celui de la proposition principale qui est de '
          {20}       ,'Cette information est obligatoire : ' // n'est utilisé que pour cegid
          {21}       ,'Ce projet n''existe pas'
          {22}       ,'Ce concurrent n''existe pas'
          {23}       ,'L''estimation réussite doit être au moins égale au maximum de celle des variantes qui est de '
          {24}       ,'L''estimation réussite doit être inférieure ou égale à celle de la proposition principale qui est de '
          {25}       ,'La saisie du champ suivant est obligatoire : '
          {26}       ,'Le projet n''est pas affecté à ce tiers. Vous devez saisir le tiers dans la fiche projet'
          );
Const Col_Tiers=1;


implementation
Uses
    TiersUtil,UtilPGI,AGlInit,EntRT,UtilRT,EntGC,UtilConfid,UtilPerspective,wCommuns
{$IFDEF AFFAIRE}
    ,COnfidentAffaire,AffaireUtil
{$ENDIF}
{$IFDEF VER150}
    ,Variants
{$ENDIF VER150}
{$IFNDEF EAGLSERVER}
  ,UtilAlertes,YAlertesConst,EntPgi
{$ENDIF EAGLSERVER}
	,facture,BTPUtil,HrichOle
  ;
{ TOM_PERSPECTIVES }

procedure TOM_PERSPECTIVES.OnChangeField(F: TField);
var Q : TQuery;
    Select : String;
begin
Inherited;
   if (F.FieldName = 'RPE_TIERS') and (GetField('RPE_TIERS')<>'') then
      begin
      Select :='SELECT T_AUXILIAIRE,T_REPRESENTANT,T_NATUREAUXI FROM TIERS WHERE T_TIERS="'+GetField('RPE_TIERS')+'"';
      if CtxAffaire in V_PGI.PGIContexte then // on ajoute ressource libre tiers
             Select :='SELECT T_AUXILIAIRE,T_REPRESENTANT,T_NATUREAUXI,YTC_RESSOURCE1,YTC_RESSOURCE2,YTC_RESSOURCE3 FROM TIERS '
             + 'lEFT JOIN TIERSCOMPL ON YTC_TIERS=T_TIERS WHERE T_TIERS="'+GetField('RPE_TIERS')+'"';
      Q := OpenSQL(Select, True);
      if not Q.Eof then
         begin
         if (DS.State = dsInsert) then
             begin
             SetField('RPE_AUXILIAIRE',Q.FindField('T_AUXILIAIRE').asstring);
             if GetField('RPE_REPRESENTANT')='' then
                  SetField('RPE_REPRESENTANT',Q.FindField('T_REPRESENTANT').asstring);
             if GetControlText('RPE_INTERVENANT')='' then
                 begin
                 if CtxAffaire in V_PGI.PGIContexte then
                    begin //mcd 25/01/2005   ajout ressource libre tiers
                    if ( GetParamsocSecur('SO_RTPERRESP','UTI') = 'RE1' ) then SetField('RPE_INTERVENANT',Q.FindField('YTC_RESSOURCE1').asstring);
                    if ( GetParamsocSecur('SO_RTPERRESP','UTI') = 'RE2' ) then SetField('RPE_INTERVENANT',Q.FindField('YTC_RESSOURCE2').asstring);
                    if ( GetParamsocSecur('SO_RTPERRESP','UTI') = 'RE3' ) then SetField('RPE_INTERVENANT',Q.FindField('YTC_RESSOURCE3').asstring);
                    end;
                 end;
             end;
         SetControlText ('NATUREAUXI',Q.FindField('T_NATUREAUXI').asstring);
{$IFDEF AFFAIRE}
         if ( (ctxAffaire in V_PGI.PGIContexte) or (ctxGCAFF in V_PGI.PGIContexte) ) and
             ( VH_GC.GASeria=True ) then
            if (pos (Q.FindField('T_NATUREAUXI').asstring,FabriqueWhereNatureAuxiAff('PRO')) = 0) then
              begin
              SetControlVisible('PAFFAIRE',false);
              SetControlVisible('BAFFAIRE',false);
              end
            else
              begin
              SetControlVisible('PAFFAIRE',true);
              SetControlVisible('BAFFAIRE',true);
              end;
{$ENDIF}              
         end
      else
        begin
        SetFocusControl('RPE_TIERS');
        LastError := 2;
        PGIBox(TraduireMemoire(TexteMessage[LastError]+' : '+GetField('RPE_TIERS')),'Code tiers');
        SetField('RPE_TIERS','');
        end;
      Ferme(Q);
      end;
end;

procedure TOM_PERSPECTIVES.OnLoadAlerte;
begin
  {$IFNDEF EAGLSERVER}
  { GC/GRC : MNG / gestion des alertes }
  if (not V_Pgi.SilentMode) and (not AfterInserting)
    and ( not ModifLot ) and (AlerteActive (TableToPrefixe(TableName))) then
      ExecuteAlerteLoad(TForm(Ecran),true);
  { GC/GRC : MNG / gestion des alertes }
  {$ENDIF !EAGLSERVER}
end;

procedure TOM_PERSPECTIVES.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var ventes : double;
    i : integer;
begin
    case Key of
      VK_F6 : {Infos compl.} if (ssAlt in Shift) then
         if GetParamSocSecur('SO_RTGESTINFOS00V',False) then RTPropoAppelParamCL(false) ;
      VK_DELETE :
        BEGIN
      if (Screen.ActiveControl = GP) and (Shift=[ssCtrl]) and ( GP.Cells[1,GP.Row] <> '' ) then
           begin
           if (GetParamSocSecur('SO_RTPROPDEVISSSPROP',True)=False) and (GP.RowCount <= 2) then
             begin
             LastError := 8;
             PGIBox(TraduireMemoire(TexteMessage[LastError]),'Liste des pièces');
      end else
             begin
        ExecuteSQl('UPDATE PIECE set GP_PERSPECTIVE=0 Where '+
        					'GP_NATUREPIECEG="'+GP.Cells[7,GP.row]+'" AND '+
                  'GP_SOUCHE="'+GP.Cells[5,GP.row]+'" AND '+
                  'GP_NUMERO='+GP.Cells[1,GP.row]+' AND '+
                  'GP_INDICEG='+GP.Cells[6,GP.row] );
        if GP.RowCount = 2 then GP.videpile (False)
                           else GP.DeleteRow (GP.Row);

             if (GetParamsocSecur('SO_RTPROPOPTCALCUL','SLP') = 'NSA') or
               (GetParamsocSecur('SO_RTPROPOPTCALCUL','SLP') = 'SSP') then
                begin
                Ventes:=0;
                for i:=1 to GP.RowCount-1 do
                    if GP.Cells[3,i] <> '' then
                       // mng 30-05-02 Ventes:=Ventes+StrToFloat(GP.Cells[3,i]) ;
                       Ventes:=Ventes+Valeur(GP.Cells[3,i]) ;
                SetField ('RPE_MONTANTPER',Ventes);
                if (ventes = 0) and (GetParamsocSecur('SO_RTPROPOPTCALCUL','SLP') = 'SSP') then
                   SetControlEnabled('RPE_MONTANTPER',True);
                end;
             end
      end else if (Screen.ActiveControl = GA) and (Shift=[ssCtrl]) and
                ( GA.Cells[1,GA.Row] <> '' ) then
                 begin
        ExecuteSQl('UPDATE AFFAIRE set AFF_PERSPECTIVE=0 Where AFF_AFFAIRE="'+GA.Cells[1,GA.row]+'"');
        if ds.State = dsBrowse then ds.Edit;
                 if GA.RowCount = 2 then
                    GA.videpile (False)
                 else
                    GA.DeleteRow (GA.Row);
      end else if ((Screen.ActiveControl = GS) and (Shift=[ssCtrl])) then
                 begin
                 if ds.State = dsBrowse then ds.Edit;
                 if GS.RowCount = 2 then
                    GS.videpile (False)
                 else
                    GS.DeleteRow (GS.Row);
                 end;
        END;
    {GED}
  VK_F3 :
       begin
       if (ssAlt in Shift) then
                       if GetParamSocSecur('SO_RTGESTIONGED',False) = True then
                         begin
                            if (ds<>nil) and not(DS.State in [dsInsert]) then
                              begin
                              Key:=0 ;
                              RTAppelGED_OnClick (Sender);
                              end;
                         end;
       end;
    81 : {Ctrl + Q - Création d'1 alerte} if (ssCtrl in Shift) then
          begin
          Key:=0 ;
          Alerte_OnClick_Rpe(Sender);
          end;
    85 : {Ctrl + U - liste des alertes du tiers} if (ssCtrl in Shift) then
          begin
          Key:=0 ;
          ListAlerte_OnClick_Rpe(Sender);
          end;
   end;
end;

procedure TOM_PERSPECTIVES.OnArgument (Arguments : String );
var Critere,ChampMul,ValMul,stAction,Libelles: string;
    x : integer;
    Arow,Acol : Integer;
    Cancel : Boolean;
begin
inherited ;
  AppliqueFontDefaut (THRichEditOle(GetControl('RPE_BLOCNOTE')));
	PieceBTP := GetParamSocSecur('SO_BTCHOIXPIECEPROPOS','DBT');
  PieceNEG := GetParamSocSecur('SO_BTCHOIXPIECENEG','DE');
  Libelles := RechDom('GCNATUREPIECEG',PieceBTP,false)+'/'+ RechDom('GCNATUREPIECEG',PieceNEG,false);
  THlabel(getControl('TGB_PIECES')).Caption := 'Liste des '+Libelles;
  TToolBarButton97(getcontrol('BPIECE')).Hint := 'Nouveau(elle) '+ Libelles;
  { mng_fq012;10859 }
  AsInsert := false;
  OldPrincipale:=0; { mng_090608_fq01210685 }
  ProchaineAction:=False;
  SaisieProposition:=False;
  MontantProposition:=0;
  { mng 25-11-2003 }
  stArgForNextAction:='';

  // mng : pour tester l'eagl
  SetControlEnabled('RPE_PERSPECTIVE',FALSE);

  if GetParamsocSecur('SO_RTPROPOPTCALCUL','SLP') = 'NSA' then
     SetControlEnabled('RPE_MONTANTPER',FALSE)
  else
     SetControlEnabled('RPE_MONTANTPER',TRUE) ;

  if GetParamsocSecur('SO_RTPROJGESTION',False) = False then
     SetControlEnabled('RPE_PROJET',FALSE);

  x := pos('NOCHANGEPROSPECT',Arguments);
  if x <> 0 then begin NoChangeProspect:=True; SetControlEnabled('RPE_TIERS',False) ; end else  NoChangeProspect:=False;
  x := pos('FICHECONTACT',Arguments);
  if x <> 0 then NoChangeContact:=True else  NoChangeContact:=False;
  x := pos('NOZOOMDEVIS',Arguments);
  if x <> 0 then NoZoomDevis:=True else  NoZoomDevis:=False;

  stTiers := '';
  stProjet := '';
  stIntervenant := '';
  stAction := '';
  iNumeroContact := 0;
  stAffaire := '';
  x := pos('MODIFLOT',Arguments);
  ModifLot := x<>0;
  if ModifLot then
    begin
    TFfiche(Ecran).MonoFiche:=true;
    StSQL := copy(Arguments,x+9,length(Arguments));
    end
  else
    begin
    Repeat
        Critere:=uppercase(ReadTokenSt(Arguments)) ;
        if Critere<>'' then
            begin
            x:=pos('=',Critere);
            if x<>0 then
               begin
               ChampMul:=copy(Critere,1,x-1);
               ValMul:=copy(Critere,x+1,length(Critere));
               if ChampMul='RPE_TIERS' then stTiers := ValMul
               //mng 25-10-00
               else if ChampMul='RPE_PROJET' then stProjet := ValMul
               else if ChampMul='RPE_INTERVENANT' then stIntervenant := ValMul
               else if ChampMul='RPE_NUMEROACTION' then iNumAction := strToInt(ValMul)
               else if ChampMul='ACTION' then stAction := ValMul
               else if ChampMul='MONTANTPROPAL' then MontantProposition := valeur(ValMul)
               else if ChampMul='RPE_NUMEROCONTACT' then iNumeroContact := strToInt(ValMul)
               else if ChampMul='AFFAIRE' then stAffaire := ValMul
               ;
               end;
            end;
    until  Critere='';
    end;
    Action:=taModif;
    if StAction='CONSULTATION' then BEGIN Action:=taConsult ; END ;
    if StAction='MODIFICATION' then BEGIN Action:=taModif ; END ;
    if StAction='CREATION' then BEGIN Action:=taCreat ; END ;

    LesColonnes :='FIXED;T_TIERS;T_LIBELLE;T_CODEPOSTAL;T_VILLE;';
    GS:=THGRID(GetControl('GCONCURRENTS'));
    GS.OnRowEnter:=GSRowEnter ;
    GS.OnRowExit:=GSRowExit ;
    GS.OnCellExit:=GSCellExit ;
    GS.OnCellEnter:=GSCellEnter ;
    GS.OnExit:=GSExit ;
    GS.onEnter := GSEnter;
    GS.OnDblClick:=GSLigneDClick ;
    GS.OnElipsisClick:=GSElipsisClick  ;
    GS.PostDrawCell:= DessineCellGS;
    GS.ColWidths[0]:=4;
    GS.ColWidths[1]:=40;
    GS.ColWidths[2]:=65;
    GS.ColWidths[3]:=20;
    GS.ColWidths[4]:=50;
    TFFiche(Ecran).Hmtrad.ResizeGridColumns(GS) ;
    TFFiche(Ecran).OnKeyDown:=FormKeyDown ;
    Cancel := false;
    Arow := 1;
    Acol := 1;
    GSRowEnter(Self,Arow,cancel,false);
    GSCellEnter(Self,Acol,Arow,cancel);
    GS.Enabled :=(Action<>taConsult);

    LesColonnesPiece :='FIXED;GP_NUMERO;GP_DATEPIECE;GP_TOTALHT;GP_DEVISE;GP_SOUCHE;GP_INDICEG;GP_NATUREPIECEG;';
    GP:=THGRID(GetControl('GPIECES'));
    GP.OnRowEnter:=GPRowEnter ;
    GP.OnRowExit:=GPRowExit ;
    GP.OnDblClick:=GPLigneDClick ;

    GP.PostDrawCell:= DessineCellGP;
    GP.ColCount := 8;
    GP.ColWidths[0]:=4;
    GP.ColWidths[1]:=50; GP.ColFormats[1]:='0';
    GP.ColWidths[2]:=50; GP.ColAligns[2]:=tacenter;
    GP.ColWidths[3]:=50; GP.ColFormats[3]:='###0.00'; GP.ColAligns[3]:=taRightJustify ;
    GP.ColWidths[4]:=50; GP.ColFormats[4]:='CB=TTDEVISE';

    GP.ColWidths[5]:=0;
    GP.ColWidths[6]:=0;
    GP.ColWidths[7]:=0;
    GP.options:=GP.Options-[goEditing] ;
    Cancel := false;
    Arow := 1;
    GPRowEnter(Self,Arow,cancel,false);
    TFFiche(Ecran).Hmtrad.ResizeGridColumns(GP) ;
{$IFDEF AFFAIRE}
  if ( (ctxAffaire in V_PGI.PGIContexte) or (ctxGCAFF in V_PGI.PGIContexte) ) and
     ( VH_GC.GASeria=True ) then
    begin
    LesColonnesAffaire :='FIXED;AFF_AFFAIRE;AFF_LIBELLE;AFF_DATEDEBUT;AFF_ETATAFFAIRE;AFF_TOTALHT';
    GA:=THGRID(GetControl('GAFFAIRES'));
    GA.OnRowEnter:=GARowEnter ;
    GA.OnRowExit:=GARowExit ;
    GA.OnDblClick:=GALigneDClick ;

    GA.PostDrawCell:= DessineCellGA;
    GA.ColWidths[0]:=4;
    GA.ColWidths[1]:=50;
    GA.ColWidths[2]:=100;
    GA.ColWidths[3]:=50; GA.ColAligns[3]:=tacenter;
    GA.ColWidths[4]:=50;
    GA.ColFormats[4]:='CB=AFETATAFFAIRE';

    GA.ColWidths[5]:=50; GA.ColFormats[5]:='###0.00'; GA.ColAligns[5]:=taRightJustify ;
    GA.options:=GA.Options-[goEditing] ;

    Cancel := false;
    Arow := 1;
    GARowEnter(Self,Arow,cancel,false);
    TFFiche(Ecran).Hmtrad.ResizeGridColumns(GA) ;
    SetControlEnabled('BAFFAIRE',(Action<>taConsult));
    end
    else
{$ENDIF}
      begin
      SetControlVisible('PAFFAIRE',false);
      SetControlVisible('BAFFAIRE',false);
      end;
    GereIsoflex;
{$ifdef GIGI}
 SetControlVisible('RPE_REPRESENTANT',false);
 SetControlVisible('TRPE_REPRESENTANT',false);
 SetControlText('TRPE_CONTACT',TraduireMemoire('Contact'));
 SetControlProperty('PAFFAIRE','caption',TraduireMemoire('Proposition de missions'));
 SetControlVisible('TGB_PIECES',false); //on cache tout ce qui concerne les devis non Géré en GI
 SetControlVisible('GPIECES',false);
 SetControlVisible('BPIECE',false);
 SetControlVisible('BVENTE',false);
 If (Not VH_GC.GaSeria)  then
   begin
   SetControlVisible('PAFFAIRE',false);
   SetControlVisible('BAFFAIRE',false);
   end;
 If (Not GetParamSocSecur ('SO_RTPROJGESTION',False)) then
   begin
    SetControlVisible('RPE_PROJET',false);
    SetControlVisible('TRPE_PROJET',false);
    SetControlVisible('TTRPE_PROJET',false);
   end;
 if (GetControl('RPE_OPERATION') <> nil) and Not GetParamsocSecur('SO_AFRTOPERATIONS',False)
    then  begin
    SetControlVisible('RPE_OPERATION',false);
    SetControlVisible('TRPE_OPERATION',false);
    end;
 if (GetControl('TGB_COMPLEMENTS') <> nil) and Not GetParamsocSecur('SO_AFRTCONCURRENT',False)
    then  begin
    SetControlVisible('TGB_COMPLEMENTS',false);
    SetControlVisible('GCONCURRENTS',false);
    end;
{$ENDIF}
if GetParamSocSecur('SO_RTGESTIONGED',False) = False then AddRemoveItemFromPopup ('POPCOMPLEMENT','MNSGEDGRC',False);
if Assigned(GetControl('mnSGEDGRC')) then
   TMenuItem(GetControl('mnSGEDGRC')).OnClick := RTAppelGED_OnClick;
if Assigned(GetControl('BDOCGEDEXIST')) then
   TToolbarButton97(GetControl('BDOCGEDEXIST')).OnClick := RTAppelGED_OnClick;
{$IFDEF GRCLIGHT}
  if not GetParamsocSecur('SO_CRMACCOMPAGNEMENT',False) then
    begin
    SetControlVisible('RPE_OPERATION',false);
    SetControlVisible('TRPE_OPERATION',false);
    SetControlVisible('TTRPE_OPERATION',false);
    SetControlVisible('RPE_PROJET',false);
    SetControlVisible('TRPE_PROJET',false);
    SetControlVisible('TTRPE_PROJET',false);
    end;
{$ENDIF GRCLIGHT}

// Pl le 19/05/07 : gestion des champs libres seulement pour KPMG pour l'instant
// Pl le 14/08/07 : FQ10622 on veut tous les nouveaux champs libre maintenant
//  if (GetParamSocSecur ('SO_AFCLIENT', 0) = 8 ) then
    GCMAJChampLibre(TForm(Ecran), False, 'COMBO', 'RPE_RPETABLELIBRE', 5, '_',true);
//  else
//    SetControlVisible('PTABLESLIBRES', false);
{$IFNDEF EAGLSERVER}
  if Assigned(GetControl('MnAlerte')) then
    if AlerteActive('RPE') then
      TMenuItem(GetControl('MnAlerte')).OnClick := Alerte_OnClick_RPE
    else
      TMenuItem(GetControl('MnAlerte')).visible:=false;

  if Assigned(GetControl('MnListAlerte')) then
    if AlerteActive('RPE') then
      TMenuItem(GetControl('MnListAlerte')).OnClick := ListAlerte_OnClick_RPE
    else
      TMenuItem(GetControl('MnListAlerte')).visible:=false;

  if Assigned(GetControl('MnGestAlerte')) and Assigned(GetControl('MnAlerte'))
     and Assigned(GetControl('MnListAlerte')) then
         TMenuItem(GetControl('MnGestAlerte')).visible := (TMenuItem(GetControl('MnAlerte')).visible)
          and (TMenuItem(GetControl('MnListAlerte')).visible);
{$ENDIF EAGLSERVER}
  { on cache les groupe box si aucun champ libre n'est géré }
  if Assigned(GetControl('CODIF3')) and Assigned(GetControl('RPE_TABLELIBREPER1'))
    and Assigned(GetControl('RPE_TABLELIBREPER2')) and Assigned(GetControl('RPE_TABLELIBREPER3')) then
      TGroupBox(GetControl('CODIF3')).Visible := TGroupBox(GetControl('RPE_TABLELIBREPER1')).Visible or
      TGroupBox(GetControl('RPE_TABLELIBREPER2')).Visible or TGroupBox(GetControl('RPE_TABLELIBREPER3')).Visible;

  if Assigned(GetControl('CODIF6')) and Assigned(GetControl('RPE_RPETABLELIBRE1'))
    and Assigned(GetControl('RPE_RPETABLELIBRE2')) and Assigned(GetControl('RPE_RPETABLELIBRE3')) then
      TGroupBox(GetControl('CODIF6')).Visible := TGroupBox(GetControl('RPE_RPETABLELIBRE1')).Visible or
      TGroupBox(GetControl('RPE_RPETABLELIBRE2')).Visible or TGroupBox(GetControl('RPE_RPETABLELIBRE3')).Visible or
      TGroupBox(GetControl('RPE_RPETABLELIBRE4')).Visible or TGroupBox(GetControl('RPE_RPETABLELIBRE5')).Visible;
  if Assigned(GetControl('BIMPRIMEFICHE')) then
    TToolbarButton97(GetControl('BIMPRIMEFICHE')).OnClick:= BIMPRIMEFICHE_OnClick;
end;

procedure TOM_PERSPECTIVES.OnNewRecord;
begin
  inherited;
  Bouge:=False;
  { mng_fq012;10859 }
  AsInsert := true;
  Action:=TaCreat;
  ProchaineAction:= GetParamsocSecur('SO_RTPROPCHAINEACTION',True);
  NouvProp := True;
  if (stTiers) <> '' then
     begin
     Setfield ('RPE_TIERS', stTiers);
     //SetField('RPE_AUXILIAIRE',TiersAuxiliaire ((stTiers), False));
     {Select :='SELECT T_AUXILIAIRE,T_REPRESENTANT,T_NATUREAUXI FROM TIERS WHERE T_TIERS="'+stTiers+'"';
     Q := OpenSQL(Select, True);
     if not Q.Eof then
        begin
        SetField('RPE_AUXILIAIRE',Q.FindField('T_AUXILIAIRE').asstring);
        SetField('RPE_REPRESENTANT',Q.FindField('T_REPRESENTANT').asstring);
        SetControlText ('NATUREAUXI',Q.FindField('T_NATUREAUXI').asstring);
        end;
     Ferme(Q);}
     end;

   //SetControlEnabled('RPE_PERSPECTIVE',False) ;
  // mng 25-10-00
  SetField('RPE_PROJET',stProjet);

  //Setfocuscontrol('RPE_TYPEPERSPECTIV');

  SetField('RPE_ETATPER',Encours);
  SetField('RPE_INTERVENANT',stIntervenant);

 SetField('RPE_DATEREALISE', V_PGI.DateEntree);
  SetField('RPE_NUMEROACTION',iNumAction);
  if ( GetParamsocSecur('SO_RTPERRESP','UTI') = 'UTI' ) then
    SetField('RPE_INTERVENANT',VH_RT.RTResponsable);

  GP.videpile (False);
  GS.videpile (False);

  if MontantProposition <> 0 then
     SetField('RPE_MONTANTPER',MontantProposition);
  SetControlText ('LECONTACT','');
  if iNumeroContact <> 0 then SetField('RPE_NUMEROCONTACT',iNumeroContact);

end;


procedure TOM_PERSPECTIVES.OnClose ;
var i,j : integer;
    Trouve : Boolean;
begin
// si annulation saisie, on maj les pièces telles qu'à l'origine
if not SaisieProposition then
begin
    // on remet celles qui ont été supprimées
    for i:=0 to TobPieceOrigine.detail.count-1 do
        begin
        Trouve:=false;
        for j:=1 to (GP.RowCount-1) do
           begin
           if (TobPieceOrigine.detail[i].GetValue('GP_SOUCHE') = GP.Cells[5,j]) and
              (TobPieceOrigine.detail[i].GetValue('GP_NUMERO') = GP.Cells[1,j]) and
              (TobPieceOrigine.detail[i].GetValue('GP_INDICEG') = GP.Cells[6,j]) then
              begin
              Trouve:=True;
              break;
              end;
           end;
           // pièce liée à l'origine a été enlevée : on la remet
           if not Trouve then
              begin
              ExecuteSQl('UPDATE PIECE set GP_PERSPECTIVE='+IntToStr(getfield('RPE_PERSPECTIVE'))+' '+
              					 'WHERE GP_NATUREPIECEG="'+TobPieceOrigine.detail[i].GetValue('GP_NATUREPIECEG')+'" AND '+
                         'GP_SOUCHE="'+TobPieceOrigine.detail[i].GetValue('GP_SOUCHE')+'" AND '+
                         'GP_NUMERO='+IntToStr(TobPieceOrigine.detail[i].GetValue('GP_NUMERO'))+' AND '+
                         'GP_INDICEG='+IntToStr(TobPieceOrigine.detail[i].GetValue('GP_INDICEG')) );
              end;
        end;
    // on enlève celle qui ont été rajoutées
    for j:=1 to (GP.RowCount-1) do
        begin
        Trouve:=false;
        for i:=0 to TobPieceOrigine.detail.count-1 do
           begin
           if (TobPieceOrigine.detail[i].GetValue('GP_SOUCHE') = GP.Cells[5,j]) and
              (TobPieceOrigine.detail[i].GetValue('GP_NUMERO') = GP.Cells[1,j]) and
              (TobPieceOrigine.detail[i].GetValue('GP_NATUREPIECEG') = GP.Cells[7,j]) and
              (TobPieceOrigine.detail[i].GetValue('GP_INDICEG') = GP.Cells[6,j]) then
              begin
              Trouve:=True;
              break;
              end;
           end;
           // pièce non liée à l'origine : on enlève le lien
           if ( not Trouve ) and (GP.Cells[1,j] <> '' ) then
              begin
              ExecuteSQl('UPDATE PIECE SET GP_PERSPECTIVE=0 WHERE '+
                        'GP_NATUREPIECEG="'+GP.Cells[7,j]+'" AND '+
                        'GP_SOUCHE="'+GP.Cells[5,j]+'" AND '+
                        'GP_NUMERO='+GP.Cells[1,j]+' AND '+
                        'GP_INDICEG='+GP.Cells[6,j] );
              end;
        end;
    { idem pour propositions d'affaires }
    if ( (ctxAffaire in V_PGI.PGIContexte) or (ctxGCAFF in V_PGI.PGIContexte) ) and
       ( VH_GC.GASeria=True ) then
      begin
      for i:=0 to TobAffaire.detail.count-1 do
          begin
          Trouve:=false;
          for j:=1 to (GA.RowCount-1) do
             begin
             if (TobAffaire.detail[i].GetValue('AFF_AFFAIRE') = GA.Cells[1,j]) then
                begin
                Trouve:=True;
                break;
                end;
             end;
             // pièce liée à l'origine a été enlevée : on la remet
          if not Trouve then
             ExecuteSQl('UPDATE AFFAIRE set AFF_PERSPECTIVE='+IntToStr(getfield('RPE_PERSPECTIVE'))
               +' Where AFF_AFFAIRE="'+TobAffaire.detail[i].GetValue('AFF_AFFAIRE')+'"');
          end;
      for j:=1 to (GA.RowCount-1) do
          begin
          Trouve:=false;
          for i:=0 to TobAffaire.detail.count-1 do
             begin
             if (TobAffaire.detail[i].GetValue('AFF_AFFAIRE') = GA.Cells[1,j]) then
                begin
                Trouve:=True;
                break;
                end;
             end;
             // pièce non liée à l'origine : on enlève le lien
             if ( not Trouve ) and (GA.Cells[1,j] <> '' ) then
                begin
                ExecuteSQl('UPDATE AFFAIRE set AFF_PERSPECTIVE=0 Where AFF_AFFAIRE="'
                  +GA.Cells[1,j]+'"');
                end;
          end;
      end;
end;

if ProchaineAction and (stArgForNextAction<>'') then
  if (VH_RT.RTCreatActions) or ( bDroitModifTiers) then
    AGLLanceFiche('RT','RTACTIONS','','','ACTION=CREATION;MONOFICHE;NOCHANGEPROSPECT;'+stArgForNextAction);

if TOBPersp<> Nil then
   TOBPersp.free;
if TobPieceOrigine<> Nil then
   TobPieceOrigine.free;
if ( (ctxAffaire in V_PGI.PGIContexte) or (ctxGCAFF in V_PGI.PGIContexte) ) and
   ( VH_GC.GASeria=True ) then
  if TobAffaire<> Nil then
     TobAffaire.free;

inherited;
end;

procedure TOM_PERSPECTIVES.OnLoadRecord;
var i,PerspVariante :integer;
    StEtatPer :string;
    QQ : TQuery;
    TobInitialeConcurrents : Tob;
    SQl : string;
begin
inherited ;
  Bouge:=False;
  //stArgForNextAction:='';
  bDroitModifTiers:=RTDroitModifTiers(GetField('RPE_TIERS'));
  for i:=1 to 3 do
    SetControlCaption('TRPE_TABLELIBREPER'+IntToStr(i),'&'+RechDom('RTLIBCHAMPSLIBRES','PL'+IntToStr(i),FALSE)) ;

  if not (DS.State in [dsInsert]) then
     begin
     StEtatPer := getfield ('RPE_ETATPER');
     if ((StEtatPer='PER') or (StEtatPer='ABA') or (StEtatPer='SIG')) then
       begin
       BTermine := True;
       SetControlCaption('TRPE_DATEFINVIE',RechDom('RTETATPERSPECTIVE',StEtatPer,FALSE)+' le');
       end else
       BTermine := False;
     end;

    SetControlEnabled('BPIECE',(Action<>taConsult));
    SetControlEnabled('BVARIANTE',(Action<>taConsult));
    SetControlEnabled('BVENTE',(Action<>taConsult));

    SetControlEnabled('RPE_TIERS',((DS.State in [dsInsert]) and not NoChangeProspect )) ;
    if GetParamsocSecur('SO_RTPROJGESTION',False) = True then
        //SetControlEnabled('RPE_PROJET',(DS.State in [dsInsert])) ;
        SetControlEnabled('RPE_PROJET',true) ;
    //SetControlEnabled('RPE_LIBELLE',(DS.State in [dsInsert])) ;
    //SetControlEnabled('RPE_TYPEPERSPECTIV',(DS.State in [dsInsert])) ;
    //SetControlEnabled('RPE_OPERATION',(DS.State in [dsInsert])) ;
    SetControlVisible('RPE_DATEFINVIE',BTermine) ;
    SetControlVisible('TRPE_DATEFINVIE',BTermine) ;


    if (not BVariante) then
    begin
      TOBPersp.free;TOBPersp:=Nil;
      TOBPersp := TOB.Create('PERSPECTIVES', nil, -1);     // pour enregistrement dans l'historique des modifs
      TOBPersp.SelectDB ('',TFFiche(Ecran).QFiche, TRUE) ;
      PerspVariante := getfield ('RPE_VARIANTE');
      PerspPrincipale := getfield ('RPE_PERSPECTIVE');

      Bprincipale := (boolean(PerspVariante=0) or boolean(PerspPrincipale=PerspVariante));
      SetControlChecked ('PROPOPRINCIPALE',Bprincipale) ;
      SetControlVisible ('RPE_VARIANTE',not Bprincipale) ;
      SetControlenabled('PROPOPRINCIPALE',not Bprincipale) ;
      PiecePerspective ('CONSULTATION'); // Affiche la piéce correspondante Devis N°3 du ....
    end else
    begin
      Bprincipale := false;
      SetControlChecked ('PROPOPRINCIPALE',False) ;
      SetControlenabled('PROPOPRINCIPALE',true) ;
      SetControlVisible ('RPE_VARIANTE',True) ;
      SetControlProperty('RPE_TIERS','TabStop',False) ;
      Setfocuscontrol('RPE_TYPEPERSPECTIV');
    end;

    ContactNom();
    if NoChangeContact = True then
       SetControlEnabled('LECONTACT',False) ;

TobPieceOrigine.free; TobPieceOrigine:=Nil;
TobPieceOrigine:=tob.create('_PIECE',Nil,-1) ;

if Action <> TaCreat then
  begin
    Sql :='SELECT GP_NATUREPIECEG,GP_NUMERO,GP_DATEPIECE,GP_TOTALHT,GP_DEVISE,GP_SOUCHE,GP_INDICEG FROM PIECE WHERE '+
          'GP_NATUREPIECEG IN ("'+PieceBTP+'","'+PieceNeg+'") AND '+
          'GP_VENTEACHAT="VEN" AND '+
          'GP_PERSPECTIVE='+IntToStr(getfield('RPE_PERSPECTIVE'))+' AND '+
          'GP_TIERS="'+getfield('RPE_TIERS')+'"';
  QQ:=OpenSQL(Sql,false,-1,'',false);
  TobPieceOrigine.LoadDetailDB('PIECE','','',QQ,false,true) ;
  ferme(QQ);
  TobPieceOrigine.PutGridDetail(GP,True,True,LesColonnesPiece,True);
  end;
if ( (ctxAffaire in V_PGI.PGIContexte) or (ctxGCAFF in V_PGI.PGIContexte) ) and
     ( VH_GC.GASeria=True ) then
  begin
  TobAffaire.free; TobAffaire:=Nil;
  TobAffaire:=tob.create('_AFFAIRE',Nil,-1) ;

  if Action <> TaCreat then
    begin
    QQ:=OpenSQL('SELECT AFF_AFFAIRE,AFF_LIBELLE,AFF_DATEDEBUT,AFF_ETATAFFAIRE,AFF_TOTALHT from AFFAIRE where AFF_AFFAIRE0="P" and AFF_PERSPECTIVE='+IntToStr(getfield('RPE_PERSPECTIVE')),false);
    TobAffaire.LoadDetailDB('AFFAIRE','','',QQ,false,true) ;
    ferme(QQ);
    TobAffaire.PutGridDetail(GA,True,true,LesColonnesAffaire,True);
    GA.Cells[3,0]:='Date début';
    GA.Cells[4,0]:='Etat';
    end;
  end;

if ( GetParamsocSecur('SO_RTPROPOPTCALCUL','SLP') = 'SSP' ) and ( GP.Cells[1,1] <> '' ) then
   SetControlEnabled('RPE_MONTANTPER',FALSE);

if (DS.State in [dsInsert]) and (GetParamSocSecur('SO_RTPROPCHAINEACTION',True)=True) then
   begin
   SetControlenabled ('BVALIDER_CREATACTION',true) ;
   SetControlVisible ('BVALIDER_CREATACTION',True) ;
   SetControlVisible ('BVALIDER',False) ;
   end
else
   begin
   SetControlVisible ('BVALIDER_CREATACTION',False) ;
   SetControlenabled ('BVALIDER',(Action<>taConsult)) ;
   SetControlVisible ('BVALIDER',True) ;
   end;

If Not(DS.State in [dsInsert]) Then
   begin
   DerniereCreate := 0;
   {Select :='SELECT T_NATUREAUXI FROM TIERS WHERE T_TIERS="'+GetField('RPE_TIERS')+'"';
   Q := OpenSQL(Select, True);
   if not Q.Eof then
      SetControlText ('NATUREAUXI',Q.FindField('T_NATUREAUXI').asstring);
   Ferme(Q);}
   end;

  // remplissage du Grid Concurrents

  TobInitialeConcurrents:=tob.create('_TIERS',Nil,-1) ;
  QQ:=OpenSQL('SELECT t_tiers,t_libelle,T_CODEPOSTAL,T_VILLE from perspectivestiers left join tiers on t_tiers=rpt_tiers where rpt_perspective='+IntToStr(getfield('RPE_PERSPECTIVE')),true);
  TobInitialeConcurrents.LoadDetailDB('TIERS','','',QQ,true,true) ;
  ferme(QQ);
  TobInitialeConcurrents.PutGridDetail(GS,True,True,LesColonnes,False);
  TobInitialeConcurrents.Free;
  ListeInitConc:='';
  for i := 0 to GS.RowCount-1 do
      if GS.Cells[1,i] <> '' then ListeInitConc:=ListeInitConc+GS.Cells[1,i]+';';
{$IFNDEF CCS3}
  AppliquerConfidentialite(Ecran,'');
{$ENDIF}
{$IFDEF EAGLCLIENT}
   if Ecran is TFFiche then
      TFFiche(Ecran).QFiche.CurrentFille.PutEcran(TFFiche(Ecran)) ;
{$ENDIF}

      if (DS.State in [dsInsert]) then
         begin
         If NoChangeProspect then
            SetFocusControl('RPE_TYPEPERSPECTIV')
            else Setfocuscontrol('RPE_TIERS') ;
         end
      else
         Setfocuscontrol('RPE_REPRESENTANT');
GestionBoutonGED;
if ModifLot then
  begin
  SetArguments(StSQL);
  Setfocuscontrol('LECONTACT');
  end;
end;


procedure TOM_PERSPECTIVES.OnUpdateRecord;
var Q : TQuery;
    nb_persp,nb_indice,i ,pourcentage : Integer;
    Btermine2,OptOk : Boolean;
    StConc,valeur,Select,CodeProspect,CodeOperation,MessageCaption,stPrincipale,StPerspective, StEtatPer : string;
    StIntervenant,NomChamp : string;
{$ifndef GIGI}
    StRepresentant : string;
{$endif}
    ToTOB,TobConcurrents,TC : TOB ;
    TiersDuProjet,GroupeDuTiersPersp,GroupeDuTiersProjet : String;
begin
  inherited;
  nb_indice := 0;

  MessageCaption:=Ecran.Caption ;
  if (DS.State in [dsInsert]) then
  begin
    CodeProspect:=GetField('RPE_TIERS') ;
    if (CodeProspect='') then
    begin
      SetFocusControl('RPE_TIERS');
      LastError := 1;
      PGIBox(TraduireMemoire(TexteMessage[LastError]),MessageCaption);
      exit;
    end else
    if not VH_RT.RTCreatPropositions and not bDroitModifTiers then
      begin
      SetFocusControl('RPE_TIERS');
      LastError :=14;
      PGIBox(TraduireMemoire(TexteMessage[LastError]),'Accès en création');
      exit;
      end;
    end;
    setfield ('RPE_TYPETIERS',GetControlText('NATUREAUXI'));
    
    StIntervenant := GetField('RPE_INTERVENANT');
    If (StIntervenant='') or (Not ExisteSQL('SELECT ARS_RESSOURCE FROM RESSOURCE where ARS_RESSOURCE="'+StIntervenant+'"')) then
    begin
      SetFocusControl('RPE_INTERVENANT');
      If (StIntervenant='') then LastError := 15 else LastError := 17;
      PGIBox(TraduireMemoire(TexteMessage[LastError]),MessageCaption);
      exit;
    end ;
{$IFNDEF GIGI}      //mcd 05/01/2005  représentant pas géré en Gi...
{$IFNDEF BTP}
    StRepresentant := GetField('RPE_REPRESENTANT');
    If (StRepresentant='') or (Not ExisteSQL('SELECT GCL_COMMERCIAL FROM COMMERCIAL where GCL_COMMERCIAL="'+StRepresentant+'"')) then
    begin
      SetFocusControl('RPE_REPRESENTANT');
      If (StRepresentant='') then LastError := 16 else LastError := 18;
      PGIBox(TraduireMemoire(TexteMessage[LastError]),MessageCaption);
      exit;
    end ;
{$ENDIF}
{$ENDIF}
    If (GetField('RPE_ETATPER')='') then
    begin
      SetFocusControl('RPE_ETATPER');
      LastError := 6;
      PGIBox(TraduireMemoire(TexteMessage[LastError]),MessageCaption);
      exit;
    end
    else begin
         if ((getField('RPE_ETATPER')='SIG') OR (getField('RPE_ETATPER')='PER') )and (getField('RPE_MOTIFPERTE1')='')  then
            begin
            SetFocusControl('RPE_MOTIFPERTE1');
            LastError := 7;
            PGIBox(TraduireMemoire(TexteMessage[LastError]),MessageCaption);
            exit;
            end;
         end;
    if VH_GC.GCIfDefCEGID then
    begin
      If (GetField('RPE_TABLELIBREPER1')='') then
      begin
        SetFocusControl('RPE_TABLELIBREPER1');
        LastError := 20;
        // modif 03-07-2002
        PGIBox(TraduireMemoire(TexteMessage[LastError])+RechDom('RTLIBCHAMPSLIBRES','PL1',FALSE),MessageCaption);
        exit;
      end;
      // modif 03-07-2002
      If (GetField('RPE_TABLELIBREPER3')='') and ( GetField('RPE_ETATPER')=Perdue ) then
      begin
        SetFocusControl('RPE_TABLELIBREPER3');
        LastError := 20;
        PGIBox(TraduireMemoire(TexteMessage[LastError])+RechDom('RTLIBCHAMPSLIBRES','PL3',FALSE),MessageCaption);
        exit;
      end;
      // modif 03-07-2002
      If (GetField('RPE_COMMENTPERTE')='') and ( GetField('RPE_ETATPER')=Perdue ) then
      begin
        SetFocusControl('RPE_COMMENTPERTE');
        LastError := 20;
        PGIBox(TraduireMemoire(TexteMessage[LastError])+TraduireMemoire('commentaire de perte'),MessageCaption);
        exit;
      end;
      // modif 11-07-2002
      If (GS.Cells[1,1]='') and ( GetField('RPE_ETATPER')=Perdue ) then
      begin
        SetFocusControl(GS.Cells[1,1]);
        LastError := 20;
        PGIBox(TraduireMemoire(TexteMessage[LastError])+TraduireMemoire('concurrent'),MessageCaption);
        exit;
      end;
    end;

  if (DS.State in [dsInsert]) then
  begin
    {Select :='SELECT T_NATUREAUXI FROM TIERS WHERE T_TIERS="'+CodeProspect+'"';
    Q := OpenSQL(Select, True);
    if Q.Eof then
    begin
      Ferme(Q) ;
      SetFocusControl('RPE_TIERS');
      LastError := 2;
      PGIBox(TraduireMemoire(TexteMessage[LastError]),MessageCaption);
      exit;
    end;
    setfield ('RPE_TYPETIERS',Q.Fields[0].AsString);
    Ferme(Q) ;}


    If (GetField('RPE_LIBELLE')='') then
    begin
      SetFocusControl('RPE_LIBELLE');
      LastError := 4;
      PGIBox(TraduireMemoire(TexteMessage[LastError]),MessageCaption);
      exit;
    end;

    If (GetField('RPE_TYPEPERSPECTIV')='') then
    begin
      SetFocusControl('RPE_TYPEPERSPECTIV');
      LastError := 3;
      PGIBox(TraduireMemoire(TexteMessage[LastError]),MessageCaption);
      exit;
    end;

    CodeOperation:=GetField('RPE_OPERATION');
    If ((CodeOperation<>'') and  Not ExisteSQL('SELECT ROP_OPERATION FROM OPERATIONS WHERE ROP_FERME<>"X" AND ROP_OPERATION="'+CodeOperation+'" AND ROP_PRODUITPGI="GRC"')) then
    begin
      SetFocusControl('RPE_OPERATION');
      LastError := 5;
      PGIBox(TraduireMemoire(TexteMessage[LastError]),MessageCaption);
      exit;
    end;

    If (GetField('RPE_ETATPER')='') then
    begin
      SetFocusControl('RPE_ETATPER');
      LastError := 6;
      PGIBox(TraduireMemoire(TexteMessage[LastError]),MessageCaption);
      exit;
    end;

    Select := 'SELECT MAX(RPE_PERSPECTIVE) FROM PERSPECTIVES' ;
    Q := OpenSQL(Select, True);
    if not Q.Eof then
       begin
       nb_persp := Q.Fields[0].AsInteger;
       setfield ('RPE_PERSPECTIVE',nb_persp + 1);
       end else
       setfield ('RPE_PERSPECTIVE',1);
    Ferme(Q) ;
    { sauf si comme variante c'est une principale que je créé ..}
    if (BVariante) and (not TCheckbox(GetControl('PROPOPRINCIPALE')).Checked) then
      begin
      SetField('RPE_VARIANTE', PerspPrincipale);
      ExecuteSQL('UPDATE PERSPECTIVES SET RPE_VARIANTE=RPE_PERSPECTIVE WHERE RPE_PERSPECTIVE='+ String(getfield('RPE_VARIANTE')));
      end
    else
      PerspPrincipale :=Getfield ('RPE_PERSPECTIVE');

    BVariante := False;
    {mng 25-10-00}SetControlEnabled('RPE_TIERS',False) ;
    //SetControlEnabled('RPE_LIBELLE',False) ;
    //SetControlEnabled('RPE_TYPEPERSPECTIV',False) ;
    //SetControlEnabled('RPE_OPERATION',False) ;
    //SetControlEnabled('PROPOPRINCIPALE',True) ;
  end; {mng else
  begin} // modification
    if (NouvProp) then  // Modification de la nouvelle proposition
      begin
      TOBPersp.SelectDB ('',TFFiche(Ecran).QFiche, TRUE) ;
      NouvProp := False;
      end;

    StPerspective := IntToStr(getfield('RPE_PERSPECTIVE'));

    if not bPrincipale then
    begin
      stPrincipale := IntToStr(getfield('RPE_VARIANTE'));
      if (TCheckbox(GetControl('PROPOPRINCIPALE')).Checked)
      and (PGIAsk('    Confirmez-vous le passage en proposition principale de cette variante ?',MessageCaption)=mrYes ) then
      begin
        pourcentage := 0;
        Q := OpenSQL('SELECT RPE_POURCENTAGE FROM PERSPECTIVES WHERE RPE_PERSPECTIVE='+stPrincipale, True);
        if not Q.Eof then pourcentage := Q.Fields[0].AsInteger ;
        ferme(Q);
        if (pourcentage > GetField ('RPE_POURCENTAGE')) then
        begin
          SetFocusControl('RPE_POURCENTAGE');
          LastError := 19;
          PGIBox(TraduireMemoire(TexteMessage[LastError])+IntTostr(pourcentage)+' %',MessageCaption);
          exit;
        end;
        if Getfield ('RPE_VARIANTE') <> 0 then
          ExecuteSQL('UPDATE PERSPECTIVES SET RPE_VARIANTE='+StPerspective+' WHERE RPE_VARIANTE='+intToStr(Getfield ('RPE_VARIANTE'))+' AND RPE_PERSPECTIVE<>'+StPerspective)
        else { mng_090608_fq01210685 }
          if OldPrincipale <> 0 then
            ExecuteSQL('UPDATE PERSPECTIVES SET RPE_VARIANTE='+StPerspective+' WHERE RPE_PERSPECTIVE='+IntToStr(OldPrincipale)+' AND RPE_PERSPECTIVE<>'+StPerspective)
          ;
        setfield ('RPE_VARIANTE',StPerspective);
      end;
    end;

    StEtatPer := THDBValCombobox(GetControl('RPE_ETATPER')).Value ;
    Btermine2 := ((StEtatPer='PER') or (StEtatPer='ABA') or (StEtatPer='SIG'));
    if (not Btermine and Btermine2) then
    begin
      setfield ('RPE_DATEFINVIE',Date);
      SetControlCaption('TRPE_DATEFINVIE',RechDom('RTETATPERSPECTIVE',StEtatPer,FALSE)+' le');
    end;
    if (not Btermine2) then
        setfield ('RPE_DATEFINVIE',iDate1900);

    SetControlVisible('RPE_DATEFINVIE',Btermine2) ;
    SetControlVisible('TRPE_DATEFINVIE',Btermine2) ;

    Select := 'SELECT MAX(RPH_INDICE) FROM PERSPHISTO WHERE RPH_PERSPECTIVE='+StPerspective;
    Q := OpenSQL(Select, True);
    if not Q.Eof then
       begin
       nb_indice := Q.Fields[0].AsInteger;
       end;
    Ferme(Q) ;
  { mng }
  { sur variante, controle % principale non inférieur à % variante }
  if (GetParamSocSecur('SO_RTPROPVARIANTE',False) = True) and (Getfield('RPE_VARIANTE') <>0) and ( Not Bprincipale )
     and ( GetField('RPE_ETATPER') <> Perdue) and ( GetField('RPE_ETATPER') <> Abandonnee) then
  begin
    Q := OpenSQL('SELECT RPE_POURCENTAGE FROM PERSPECTIVES WHERE RPE_PERSPECTIVE='+
         IntToStr(Getfield('RPE_VARIANTE'))+' AND RPE_VARIANTE = RPE_PERSPECTIVE',true) ;
    if not Q.Eof then
      if Q.Fields[0].AsInteger < Getfield('RPE_POURCENTAGE') then
      begin
        SetFocusControl('RPE_POURCENTAGE');
        LastError := 24;
        PGIBox(TraduireMemoire(TexteMessage[LastError])+IntTostr(Q.Fields[0].AsInteger)+' %',MessageCaption);
        ferme(Q);
        exit;
      end;
      ferme(Q);
  end;

  if (DS.State = dsEdit) then
  begin
    { sur principale, controle % variante non supérieur à % principale }
    if (GetParamSocSecur('SO_RTPROPVARIANTE',False) = True) and ( Bprincipale )
       and ( GetField('RPE_ETATPER') <> Perdue) and ( GetField('RPE_ETATPER') <> Abandonnee) then
    begin
      Q := OpenSQL('SELECT MAX(RPE_POURCENTAGE) FROM PERSPECTIVES WHERE RPE_VARIANTE="'+StPerspective+
         '" AND RPE_POURCENTAGE>'+IntToStr(Getfield('RPE_POURCENTAGE'))+' AND RPE_VARIANTE <> RPE_PERSPECTIVE',true) ;
      if not Q.Eof then
        if Q.Fields[0].AsInteger > Getfield('RPE_POURCENTAGE') then
        begin
          SetFocusControl('RPE_POURCENTAGE');
          LastError := 23;
          PGIBox(TraduireMemoire(TexteMessage[LastError])+IntTostr(Q.Fields[0].AsInteger)+' %',MessageCaption);
          ferme(Q);
          exit;
        end;
        ferme(Q);
    end;

    if not Bouge then  // pour parer le cas du bouge ou l'on ne va pas creer d'historique
      begin
      ToTOB:=Tob.Create('PERSPHISTO', Nil, -1) ;
      PERSPECTIVESTOBCopieChamp(TOBPersp, ToTOB);
      ToTOB.PutValue('RPH_INDICE', nb_indice + 1);
      ToTOB.InsertDB(Nil);
      ToTOB.free;
      end;
  end;

if (GetField ('RPE_PROJET') <> '') and (THEdit(GetControl('RPE_PROJET')).Enabled = True ) then
   begin
   if Not ExisteSQL ('SELECT Rpj_tiers FROM projets WHERE rpj_projet="'+GetField ('RPE_projet')+'"') then
      begin
      SetFocusControl('RPE_PROJET');
      LastError := 21;
      PGIBox(TraduireMemoire(TexteMessage[LastError]),MessageCaption);
      exit;
      end;

   if not GetParamSocSecur ('SO_RTPROJMULTITIERS',True) then
       begin
{       if ( ExisteSQL('SELECT RPJ_TIERS FROM PROJETS WHERE RPJ_PROJET="'+GetField('RPE_PROJET')+
              '" AND RPJ_TIERS <> "'+GetField ('RPE_TIERS')+'" AND RPJ_TIERS <> " "') ) or
          ( ExisteSQL('SELECT RPE_TIERS FROM PERSPECTIVES WHERE RPE_PROJET="'+
           GetField ('RPE_PROJET')+'" AND RPE_TIERS <> "'+GetField ('RPE_TIERS')+'"') ) then   }
         if (Not ExisteSQL ('SELECT RPJ_PROJET FROM PROJETS WHERE (RPJ_PROJET="'+
            GetField('RPE_PROJET')+'" and RPJ_AUXILIAIRE="'+GetField('RPE_AUXILIAIRE')+'")')) then
         begin
           TiersDuProjet := '';
           Q:=OpenSql('SELECT RPJ_TIERS FROM projets WHERE RPJ_PROJET="'+GetField ('RPE_PROJET')+'"',TRUE);
           try
             if not Q.EOF then TiersDuProjet:= Q.FindField('RPJ_TIERS').asstring;
           finally
             ferme(Q);
           end;
           SetFocusControl('RPE_PROJET');
           if TiersDuProjet = '' then LastError := 26
           else LastError := 10;
           PGIBox(TraduireMemoire(TexteMessage[LastError]),MessageCaption);
           exit;
         end;
       end
   else
       if GetParamSocSecur ('SO_RTPROPGROUPE',False) then
          begin
          GroupeDuTiersPersp:='';
          TiersDuProjet:='';
          GroupeDuTiersProjet:='';
          // Tiers du projet
          Select :='SELECT Rpj_tiers FROM projets WHERE rpj_projet="'+GetField ('RPE_projet')+'"';
          Q := OpenSQL(Select, True);
          if not Q.Eof then
             TiersDuProjet:=Q.Fields[0].AsString
          else
              begin
              SetFocusControl('RPE_PROJET');
              LastError := 11;
              PGIBox(TraduireMemoire(TexteMessage[LastError]),MessageCaption);
              Ferme(Q);
              exit;
              end;
          Ferme (Q);

          // Groupe du tiers de la perspective
          Select :='SELECT T_SOCIETEGROUPE FROM TIERS WHERE T_TIERS="'+GetField ('RPE_TIERS')+'"';
          Q := OpenSQL(Select, True);
          if not Q.Eof then
             GroupeDuTiersPersp:=Q.Fields[0].AsString ;
          Ferme (Q);
          OptOk:=False;
          // Groupe du tiers du projet
          Select :='SELECT T_SOCIETEGROUPE FROM TIERS WHERE T_TIERS="'+TiersDuProjet+'"';
          Q := OpenSQL(Select, True);
          if not Q.Eof then
             GroupeDuTiersProjet:=Q.Fields[0].AsString ;
          Ferme (Q);
          if ( GroupeDuTiersPersp = TiersDuProjet ) then OptOk:=True
          else
             if ( GroupeDuTiersPersp = GroupeDuTiersProjet ) and ( GroupeDuTiersPersp <> '' ) then OptOk:=True
             else
               if ( GetField ('RPE_TIERS') =  TiersDuProjet ) then OptOk:=True
               else
                  if ( GetField ('RPE_TIERS') =  GroupeDuTiersProjet) then OptOk:=True
          ;

          if Not OptOk then
              begin
              SetFocusControl('RPE_PROJET');
              LastError := 12;
              PGIBox(TraduireMemoire(TexteMessage[LastError]),MessageCaption);
              exit;
              end;
          end ;
      end
   else
   if (GetField ('RPE_PROJET') = '') and ( GetParamSocSecur('SO_RTPROPATTACHPROJET',False) ) then
      begin
        SetFocusControl('RPE_PROJET');
        LastError := 13;
        PGIBox(TraduireMemoire(TexteMessage[LastError]),MessageCaption);
        exit;
      end;

{$IFNDEF CCS3}
NomChamp:=VerifierChampsObligatoires(Ecran,'');
if NomChamp<>'' then
      begin
      NomChamp:=ReadTokenSt(NomChamp);
      SetFocusControl(NomChamp) ;
      LastError:=25; LastErrorMsg:=TexteMessage[LastError]+champToLibelle(NomChamp);
      exit;
      end;
{$ENDIF}

{ GC/GRC : MNG / gestion des alertes }
if (ds<>nil) and (not V_Pgi.SilentMode) 
  and ( not ModifLot ) and (AlerteActive (TableToPrefixe(TableName))) then
    if (not ExecuteAlerteUpdate(TForm(Ecran),true)) then
      LastError := 99;
{ GC/GRC : MNG fin / gestion des alertes }
ListeFinConc:='';
for i := 0 to GS.RowCount-1 do
    if GS.Cells[1,i] <> '' then ListeFinConc:=ListeFinConc+GS.Cells[1,i]+';';

Trim(ListeFinConc);
Trim(ListeInitConc);
if ListeFinConc <> ListeInitConc then
    begin
    ExecuteSql ('Delete from PERSPECTIVESTIERS where RPT_PERSPECTIVE='+IntToStr(getfield('RPE_PERSPECTIVE')));
    TobConcurrents := TOB.Create ('les concurrents', Nil, -1);
    i:=1;
    valeur:='';
    Repeat
      StConc:=GS.Cells[1,i] ;
      if StConc<>'' then
         begin
         TC:=Tob.create('PERSPECTIVESTIERS',TobConcurrents,-1);
         TC.PutValue('RPT_PERSPECTIVE', getfield('RPE_PERSPECTIVE') );
         TC.PutValue('RPT_TIERS', StConc );
         end;
      Inc(i);
    until  i=21; // nb lignes maxi du grid=20
    if TobConcurrents.Detail.Count <> 0 then
       TobConcurrents.InsertDB(Nil);
    TobConcurrents.free;
    end;

if (GetField ('RPE_NUMEROACTION') <> 0) and (GetField ('RPE_NUMEROACTION') <> Null) then
    begin
    Select:='update actions set RAC_PERSPECTIVE='+IntToStr(GetField('RPE_PERSPECTIVE'))+
    ' where rac_numaction='+IntToStr(GetField('RPE_NUMEROACTION'))+
    ' and rac_tiers="'+GetField('RPE_tiers')+'"';
     ExecuteSql (Select);
    end;
// mng 02-03-01if ((GetField ('RPE_INTERVINT') <> '<<Tous>>')
// mng 02-03-01   and (GetField ('RPE_INTERVENANT') <> '')) then MajIntervintPersp();

if (DS.State in [dsInsert]) then
  begin
  if (GetField ('RPE_ETATPER') <> Perdue) and (GetField ('RPE_ETATPER') <> Signee) and
     (GetField ('RPE_ETATPER') <> Abandonnee) then
     begin
     stArgForNextAction:='RAC_TIERS='+GetField('RPE_TIERS')+';RAC_PROJET='+GetField('RPE_PROJET')+';RAC_PERSPECTIVE='+IntToStr(GetField('RPE_PERSPECTIVE'));
     if GetField('RPE_OPERATION') <> '' then
        stArgForNextAction:=stArgForNextAction+';RAC_OPERATION='+GetField('RPE_OPERATION');
     if GetParamSocSecur ('SO_RTPROPACTIONOBLIG',False) then
        stArgForNextAction:=stArgForNextAction+';ORIGINEPERSP';
     if stAffaire <> '' then
        stArgForNextAction:=stArgForNextAction+';RAC_AFFAIRE='+stAffaire+';NOCHANGEAFFAIRE';
     if NoChangeContact = True then stArgForNextAction:=stArgForNextAction+';FICHECONTACT;RAC_NUMEROCONTACT='+IntToStr(GetField('RPE_NUMEROCONTACT'))
        else if GetField('RPE_NUMEROCONTACT') <> 0 then stArgForNextAction:=stArgForNextAction+';RAC_NUMEROCONTACT='+IntToStr(GetField('RPE_NUMEROCONTACT'));
     end;
  end;

  If (ds<>nil) then
  begin
    SaisieProposition:=True;
    if (DS.State in [dsInsert])
       then DerniereCreate := GetField('RPE_PERSPECTIVE')
       else if (DerniereCreate = GetField('RPE_PERSPECTIVE')) then Ecran.Close;
  end;

end;

procedure TOM_PERSPECTIVES.OnAfterUpdateRecord;
begin
inherited;
  if (DS<>nil) and (  { mng_fq012;10859 }  AsInsert) and (GetParamSocSecur('SO_RTGESTINFOS00V',False)) then
    RTPropoAppelParamCL(true);
 AsInsert :=false;
 if ModifLot then TFFiche(ecran).BFermeClick(nil);
end;


procedure TOM_PERSPECTIVES.PerspGenereAction;
begin
inherited;
if ProchaineAction and (stArgForNextAction<>'') then
   begin
   if (VH_RT.RTCreatActions) or ( bDroitModifTiers) then
     AGLLanceFiche('RT','RTACTIONS','','','ACTION=CREATION;MONOFICHE;NOCHANGEPROSPECT;'+stArgForNextAction);
   stArgForNextAction:= '';
   ProchaineAction:=false;
   end;
end;

procedure TOM_PERSPECTIVES.NouvellePiece;
var boo : boolean;
		TypeDocs : string;
begin
boo:=false;
if Action=TaCreat then
   begin
   Boo := ProchaineAction;
   ProchaineAction:=false;
   end;

Bouge:=True; // pour ne pas creer l'historique dans ce cas
if not TFFiche(Ecran).Bouge(nbPost) then exit ;
Bouge:=False;

{if Action=TaCreat then TFFiche(Ecran).Caption:= 'Proposition créée'
   else TFFiche(Ecran).Caption:= 'Proposition enregistrée';}
   TypeDocs :=PieceNeg+','+PieceBTP+',';
  AGLLanceFiche ('RT','RTPIECETIERS_MUL','GP_TIERS='+getField('RPE_TIERS')+';GP_VIVANTE=X'+
                     ';GP_PERSPECTIVE=0;PERSPECTIVE='+intToStr(getField('RPE_PERSPECTIVE'))+
                     ';NATUREAUXI='+GetControlText('NATUREAUXI'),'','NOUVELLEPIECE');
if Boo then
  ProchaineAction:=Boo;

DS.edit; // pour reforcer la mise à jour.

ChargeGrid(False);
//if not TFFiche(Ecran).Bouge(nbPost) then exit ;
end;

procedure TOM_PERSPECTIVES.NouvelleAffaire;
var boo : boolean;
    StArgument,StChamps,tmp : string;
begin
  boo:=false;
  if Action=TaCreat then
     begin
     Boo := ProchaineAction;
     ProchaineAction:=false;
     end;

  Bouge:=True; // pour ne pas creer l'historique dans ce cas
  if not TFFiche(Ecran).Bouge(nbPost) then exit ;
  Bouge:=False;

{  if Action=TaCreat then TFFiche(Ecran).Caption:= 'Proposition créée'
     else TFFiche(Ecran).Caption:= 'Proposition enregistrée';}

  StChamps := 'AFF_STATUTAFFAIRE=PRO;AFF_TYPEAFFAIRE=NOR;AFF_TIERS='+GetField('RPE_TIERS');

  StArgument := 'STATUT:PRO;NOCHANGESTATUT;NOCHANGETIERS';

  //Stargument := Stargument + ';NOFILTRE';
  Stargument := Stargument + ';ADMINGRAYED';

  /////////////////////////
  tmp := AGLLanceFiche('RT', 'RTAFFAIRERECH_MUL', StChamps, '', StArgument);
  if tmp<>'' then
     begin
     ExecuteSQl('UPDATE AFFAIRE set AFF_PERSPECTIVE='+IntToStr(Getfield('RPE_PERSPECTIVE'))
        +' Where AFF_AFFAIRE="'+tmp+'"');
     end;
  if Boo then
    ProchaineAction:=Boo;
  if tmp<>'' then
     begin
     DS.edit;
     ChargeGridGA;
     end;
end;

function  TOM_PERSPECTIVES.GetTypePiece : string;
begin

end;

procedure TOM_PERSPECTIVES.PiecePerspective (TypeAppel:String);
var
  stCleDoc,stPiece,NumeroPiece,DatePiece:string;
  CleDoc: R_CleDoc;
  Param: R_SaisiePieceParam;
  NextPiece,TypePiece : string;
begin
    if (TypeAppel = 'CREATION') then
    begin
      FillChar(CleDoc,Sizeof(CleDoc),#0) ;
      TypePiece := GetTypePiece;
      CleDoc.NaturePiece:= TypePiece;
			CleDoc.DatePiece:=V_PGI.DateEntree ;
      CleDoc.Souche:='' ; CleDoc.NumeroPiece:=0 ; CleDoc.Indice:=0 ;
//      InitRecordSaisiePieceParam(Param);
//      SaisiePiece(CleDoc, taCreat, Param);
				SaisiePiece(CleDoc, taCreat);
    end else
    begin
      stCleDoc := getfield ('RPE_PIECE');
      //inexistant dans la fiche ? SetControlText ('TTRPE_PIECE','' );
      if ( stCleDoc <> '') then
      begin
        FillChar (CleDoc, Sizeof(CleDoc), #0);
        CleDoc.NaturePiece := ReadTokenSt(stCleDoc);
        DatePiece := ReadTokenSt(stCleDoc);
        CleDoc.Souche := ReadTokenSt(stCleDoc);
        NumeroPiece := ReadTokenSt(stCleDoc);

        StPiece := RechDom('GCNATUREPIECEG',CleDoc.NaturePiece,False) ;
        StPiece := StPiece + ' N° ' +NumeroPiece+' du '+DatePiece;
        //inexistant dans la fiche ? SetControlText ('TTRPE_PIECE',StPiece );
      end;
    end;
end;

procedure TOM_PERSPECTIVES.VariantePerspective (TypeAppel:String);
var TobForm : TOB;
    i : integer;
    StTable,StNom : String;
begin
  if (TypeAppel = 'CREATION') then
  begin
    BVariante := true;
    OldPrincipale:=GetField('RPE_PERSPECTIVE'); { mng_090608_fq01210685 }
    StTable := 'PERSPECTIVES'; //eAGL GetTableNameFromDataSet(DS);
    TobForm := TOB.Create (StTable, Nil, -1);
    TobForm.SelectDB ('',TFFiche(Ecran).QFiche, TRUE);
    TFFiche (Ecran).Bouge (NbInsert);
    TFFiche(ecran).typeaction := taCreat;
    for i := 1 to TobForm.NbChamps do
      begin
      StNom:= TobForm.GetNomChamp(i);
      SetField (TobForm.GetNomChamp(i), TobForm.GetValeur (i));
      end;
    //SetField('RPE_VARIANTE', PerspPrincipale);
    SetField('RPE_PERSPECTIVE', 0);
    TobForm.free;
  end;
end;




{****************************************************************
Auteur  ...... : AB
Créé le ...... : 13/07/2000
Modifié le ... :   /  /
Description .. : Copie les valeurs de champs d'une TOB vers une autre
Mots clefs ... : TOB;COPIE;CHAMPS
*****************************************************************}
{procedure TOM_PERSPECTIVES.TOBCopieChamp(FromTOB, ToTOB : TOB);
var iChamp , iTableLigne, i_pos: integer;
    FieldNameTo,FieldNameFrom,St:string;
begin
iTableLigne := PrefixeToNum ('RPE');
for iChamp := 1 to High(V_PGI.DEChamps[iTableLigne]) do
  begin
  FieldNameFrom := V_PGI.DEChamps[iTableLigne, iChamp].Nom ;
  St := FieldNameFrom ;
  i_pos := Pos ('_', St) ;
  Delete (St, 1, i_pos-1) ;
  FieldNameTo := 'RPH'+ St ;
  ToTOB.PutValue(FieldNameTo, FromTOB.GetValue(FieldNameFrom));
  end;
end;    }


procedure TOM_PERSPECTIVES.ClickPrincipal;
begin
   if (not BPrincipale and TCheckbox(GetControl('PROPOPRINCIPALE')).Checked) then
      DS.edit;
end;

procedure TOM_PERSPECTIVES.ChargeGrid (ForceCalcul : Boolean);
var QQ : TQuery;
Tobpiece : TOB;
Ventes : double;
i : integer;
begin
TobPiece:=tob.create('_PIECE',Nil,-1) ;
QQ:=OpenSQL('SELECT GP_NUMERO,GP_DATEPIECE,GP_TOTALHT,GP_DEVISE,GP_SOUCHE,GP_INDICEG,GP_NATUREPIECEG from PIECE WHERE '+
						'GP_NATUREPIECEG IN ("'+PieceBTP+'","'+PieceNEG+'") and '+
            'GP_VENTEACHAT="VEN" and GP_PERSPECTIVE='+IntToStr(getfield('RPE_PERSPECTIVE'))+' and '+
            'GP_TIERS="'+getfield('RPE_TIERS')+'"',true);
TobPiece.LoadDetailDB('PIECE','','',QQ,false,true) ;
ferme(QQ);
TobPiece.PutGridDetail(GP,True,True,LesColonnesPiece,True);
// calcul de la vente prévue si option
if (GetParamsocSecur('SO_RTPROPOPTCALCUL','SLP') = 'NSA') or
   (GetParamsocSecur('SO_RTPROPOPTCALCUL','SLP') = 'SSP') or ( ForceCalcul) then
    begin
    Ventes:=0;
    for i:=0 to TobPiece.detail.count-1 do
        Ventes:=Ventes+TobPiece.detail[i].GetValue('GP_TOTALHT') ;
    SetField ('RPE_MONTANTPER',Ventes);
    end;
TobPiece.free ;
if ( GetParamsocSecur('SO_RTPROPOPTCALCUL','SLP') = 'SSP' ) and ( GP.Cells[1,1] <> '' ) then
   SetControlEnabled('RPE_MONTANTPER',FALSE);
end;

procedure TOM_PERSPECTIVES.ChargeGridGA ;
var QQ : TQuery;
TobAff : TOB;
begin
  TobAff:=tob.create('_PIECE',Nil,-1) ;
  QQ:=OpenSQL('SELECT AFF_AFFAIRE,AFF_LIBELLE,AFF_DATEDEBUT,AFF_ETATAFFAIRE,AFF_TOTALHT from AFFAIRE where AFF_AFFAIRE0="P" and AFF_PERSPECTIVE='+IntToStr(getfield('RPE_PERSPECTIVE')),false);
  TobAff.LoadDetailDB('AFFAIRE','','',QQ,false,true) ;
  ferme(QQ);
  TobAff.PutGridDetail(GA,True,true,LesColonnesAffaire,True);
  GA.Cells[3,0]:='Date début';
  GA.Cells[4,0]:='Etat';
  TobAff.free ;
end;

procedure TOM_PERSPECTIVES.EnvoiMessage;
var StDestinataire,StSQL,StWhere,StOr: string;
    IContact,i : integer;
    Q : TQuery;
    TobContact:TOB;
    BODY:HTstrings;
begin
stWhere:='';
StOr:='';
BODY:=Nil;
// mng 02-03-01 : plus de destinataire StDestinataire:=GetField ('RPE_DESTMAIL');
StDestinataire:='';
if StDestinataire <> TraduireMemoire('<<Tous>>') then
   begin
   While StDestinataire <> '' do
      begin
      i :=ReadTokenI(StDestinataire);
      StWhere := StWhere + StOr + 'C_NUMEROCONTACT='+ IntToStr(i);
      StOr := ' OR ';
      end;
   if GetField ('RPE_NUMEROCONTACT') <> 0 then
      begin
      StWhere := StWhere + StOr + 'C_NUMEROCONTACT='+ IntToStr(GetField('RPE_NUMEROCONTACT'));
      end;
   end;
StSQL:='SELECT C_NUMEROCONTACT,C_NOM,C_FONCTION,C_TELEPHONE,C_RVA'+
       ',C_CIVILITE FROM CONTACT WHERE C_TYPECONTACT="T" AND C_AUXILIAIRE = "'+GetField('RPE_AUXILIAIRE')+'"';
if StWhere<>'' then StSQL:=StSQL+'AND ('+stWhere+')';

TobContact:=TOB.create ('les contacts',NIL,-1);
Q := OpenSql (StSQL,TRUE);
TobContact.LoadDetailDB ('CONTACT', '', '', Q, False);
ferme(Q) ;
StDestinataire := '';
StWhere := '';
for i:=0 to TobContact.detail.count-1 do
    begin
    if TobContact.detail[i].GetValue('C_RVA') <> '' then
       begin
       IContact:=TobContact.detail[i].GetValue('C_NUMEROCONTACT');
       if IContact = GetField('RPE_NUMEROCONTACT') then
          begin
          StWhere := TobContact.detail[i].GetValue('C_RVA');
          end
       else
           begin
           StDestinataire := StDestinataire + TobContact.detail[i].GetValue('C_RVA') +';';
           end;  
       end;
    end;
if StWhere <> '' then
   begin
   PGIEnvoiMail ('',StWhere,StDestinataire,body,'',false,1,'','');
   end
else
   begin
   PGIBox('Pas d''adresse e-mail (interlocuteur)','');
   end;
TobContact.free;
end;

procedure TOM_PERSPECTIVES.GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
GS.InvalidateRow(ou) ;
StTiersConc:=GS.Cells[1,GS.Row];
end;

procedure TOM_PERSPECTIVES.GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  GS.InvalidateRow(ou) ;
end;

procedure TOM_PERSPECTIVES.GSCellEnter(Sender: TObject; var ACol,
                                        ARow: Integer; var Cancel: Boolean);
begin
if Not Cancel then
   if (GS.Col=Col_Tiers) then
       begin
       GS.ElipsisButton:=true;
       GS.options:=GS.Options+[goEditing] ;
       GS.editormode:=True ;
       end ;
end;

procedure TOM_PERSPECTIVES.GSCellExit(Sender: TObject; var ACol,
                                        ARow: Integer; var Cancel: Boolean);
begin
  if (ACol=Col_Tiers) and ( ARow > 0 ) then
  begin
    if TestConcurrent(ACol,ARow,false) = false then
    begin
    GS.ElipsisButton:=false; GS.options:=GS.Options-[goEditing] ;GS.editormode:=False ;
    end ;
  end;
end;

procedure TOM_PERSPECTIVES.GSEnter(Sender: TObject);
var Acol,Arow : integer;
		Cancel : boolean;
begin
	Arow := GS.row;
  Acol := 1;
  Cancel := false;
  GSRowEnter(Self,Arow,Cancel,False);
  if not Cancel then
  begin
    GSCellEnter(Self,Acol,Arow,cancel);
  end;
end;

procedure TOM_PERSPECTIVES.GSExit(Sender: TObject);
var cancel: boolean;
  ACOl, ARow: integer;
begin
  ACol := GS.Col;
  ARow := GS.Row;
  Cancel := False;
  GSCellExit(Sender, ACol, ARow, Cancel);
end;

procedure TOM_PERSPECTIVES.DessineCellGS ( ACol, ARow: Longint; Canvas : TCanvas; AState: TGridDrawState);
begin
DessineCell( GS,ACol,ARow,Canvas,AState);
end;
procedure TOM_PERSPECTIVES.DessineCellGP ( ACol, ARow: Longint; Canvas : TCanvas; AState: TGridDrawState);
begin
DessineCell( GP,ACol,ARow,Canvas,AState);
end;

{$IFDEF AFFAIRE}
procedure TOM_PERSPECTIVES.DessineCellGA ( ACol, ARow: Longint; Canvas : TCanvas; AState: TGridDrawState);
begin
DessineCell( GA,ACol,ARow,Canvas,AState);
end;
{$ENDIF AFFAIRE}

procedure DessineCell (GG : Thgrid; ACol, ARow: Longint; Canvas : TCanvas; AState: TGridDrawState);
var Triangle : array[0..2] of TPoint ;
  Arect: Trect ;
begin
If Arow < GG.Fixedrows then exit ;
if (gdFixed in AState) and (ACol = 0) then
    begin
    Arect:=GG.CellRect(Acol,Arow) ;
    Canvas.Brush.Color := GG.FixedColor;
    Canvas.FillRect(ARect);
      if (ARow = GG.row) then
         BEGIN
         Canvas.Brush.Color := clBlack ;
         Canvas.Pen.Color := clBlack ;
         Triangle[1].X:=ARect.Right-2 ; Triangle[1].Y:=((ARect.Top+ARect.Bottom) div 2) ;
         Triangle[0].X:=Triangle[1].X-5 ; Triangle[0].Y:=Triangle[1].Y-5 ;
         Triangle[2].X:=Triangle[1].X-5 ; Triangle[2].Y:=Triangle[1].Y+5 ;
         if false then Canvas.PolyLine(Triangle) else Canvas.Polygon(Triangle) ;
         END ;
    end;
end;


procedure TOM_PERSPECTIVES.GSLigneDClick (Sender: TObject);
begin
  if (StTiersConc = '') then Exit;
  AGLLanceFiche('GC','GCTIERS','',TiersAuxiliaire ((StTiersConc), False),ActionToString(Action)+';T_NATUREAUXI=CON');
  //Onload;
end;

procedure TOM_PERSPECTIVES.GSElipsisClick(Sender: TObject);
begin
TestConcurrent(GS.Col,GS.Row,true);
end;

function TOM_PERSPECTIVES.TestConcurrent ( ACol,ARow : integer; Elipsis : boolean) : boolean;
var StClient,NewClient : String;
    i : Integer;
    Trouve : Boolean;
begin
Result:=False;
if ACol = Col_Tiers then
    begin
    Trouve:=False;
    StClient:=Uppercase(GS.Cells[Col_Tiers,ARow]);
    if Elipsis then
      NewClient:= AGLLanceFiche ('RT', 'RTPROSPECT_RECH','T_TIERS='+GS.Cells[Col_Tiers,ARow] , '', '')
    else
      NewClient:= StClient;
    if NewClient = '' then exit;
    if ( not(DS.State in [dsInsert,dsEdit]) ) and ( StClient <> NewClient )
       then DS.edit;
    // on controle que ce concurrent n'existe pas deja dans le grid
    for i := 1 to GS.RowCount-1 do
        if (i<>ARow) and (UpperCase(GS.Cells[Col_Tiers,i]) = UpperCase(NewClient)) {and (NewClient <> '')} then
           begin
           Trouve:=True;
           break;
           end;
    if not Trouve then
    begin
      GS.Cells[Col_Tiers,ARow]:=NewClient;
      if not AfficheTiers(ARow) then
      begin
        LastError := 22;
        PGIBox(TraduireMemoire(TexteMessage[LastError]),'Liste des concurrents');
        GS.Cells[1,ARow]:=''; GS.Cells[2,ARow]:=''; GS.Cells[3,ARow]:=''; GS.Cells[4,ARow]:='';
        GS.Row:=ARow;
        GS.Col:=Col_Tiers;
        Result:=True;
      end;
    end
    else
      begin
      LastError := 9;
      PGIBox(TraduireMemoire(TexteMessage[LastError]),'Liste des concurrents');
      GS.Cells[1,ARow]:=''; GS.Cells[2,ARow]:=''; GS.Cells[3,ARow]:=''; GS.Cells[4,ARow]:='';
      GS.Row:=ARow;
      GS.Col:=Col_Tiers;
      Result:=True;
      end;
    end;
end;

procedure TOM_PERSPECTIVES.GPRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
GP.InvalidateRow(ou) ;
end;

procedure TOM_PERSPECTIVES.GPRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
GP.InvalidateRow(ou) ;
end;
{$IFDEF AFFAIRE}
procedure TOM_PERSPECTIVES.GARowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
GA.InvalidateRow(ou) ;
end;

procedure TOM_PERSPECTIVES.GARowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
GA.InvalidateRow(ou) ;
end;
{$ENDIF}

procedure TOM_PERSPECTIVES.GPLigneDClick (Sender: TObject);
var
  CleDoc : R_CleDoc;
  SaveAutoSave : boolean ;
  Param: R_SaisiePieceParam;
begin
if NoZoomDevis then exit;
if GP.Cells[2, GP.Row] <> '' then
    begin
    // on valide d'abord sinon problème de multi à la prochaine validation
    Bouge:=True; // pour ne pas creer l'historique dans ce cas
    if not TFFiche(Ecran).Bouge(nbPost) then exit ;
    Bouge:=False;
    DecodeRefPiece ( GP.Cells[7, GP.Row]+ ';' +
    								 GP.Cells[2, GP.Row] + ';' +
                     GP.Cells[5,GP.Row] + ';' +
                     GP.CellValues[1,GP.Row] +';' +
                     GP.CellValues[6,GP.Row] + ';',
        CleDoc);
//    InitRecordSaisiePieceParam(Param);
//    SaisiePiece(CleDoc, Action, Param);
		SaisiePiece(CleDoc, Action);
    //DS.Refresh ;
    SaveAutoSave:=TFFiche(ecran).FAutoSave.checked ;
    TFFiche(ecran).FAutoSave.checked:=True;
    TFFiche(ecran).OM.RefreshDB ;
    TFFiche(ecran).FAutoSave.checked:=SaveAutoSave;
{$IFDEF EAGLCLIENT}
    ChargeGrid(False);
{$ENDIF}
    end;
end;
{$IFDEF AFFAIRE}
procedure TOM_PERSPECTIVES.GALigneDClick (Sender: TObject);
var SaveAutoSave : boolean ;
Tmp : string;
begin
if NoZoomDevis then exit;
if GA.Cells[1, GA.Row] <> '' then
    begin
    // on valide d'abord sinon problème de multi à la prochaine validation
    Bouge:=True; // pour ne pas creer l'historique dans ce cas
    if not TFFiche(Ecran).Bouge(nbPost) then exit ;
    Bouge:=False;
    if ctxscot in V_PGI.PGIContexte then tmp :='MISSION'  //mcd 06/10/2005 pas même fiche GI et GA
       else tmp := 'AFFAIRE';
    if AGLJaiLeDroitFiche(['AFFAIRE','ACTIOn=MODIFICATION','P'],3)
     then AglLanceFiche('AFF',tmp,'',GA.Cells[1, GA.Row],'ACTION=MODIFICATION;MONOFICHE')
    else
      if AGLJaiLeDroitFiche(['AFFAIRE','ACTIOn=CONSULTATION','P'],3)
       then AglLanceFiche('AFF',tmp,'',GA.Cells[1, GA.Row],'ACTION=CONSULTATION;MONOFICHE');

    //DS.Refresh ;
    SaveAutoSave:=TFFiche(ecran).FAutoSave.checked ;
    TFFiche(ecran).FAutoSave.checked:=True;
    TFFiche(ecran).OM.RefreshDB ;
    TFFiche(ecran).FAutoSave.checked:=SaveAutoSave;
{$IFDEF EAGLCLIENT}
    //ChargeGrid(False);
{$ENDIF}
    end;
end;
{$ENDIF}

function TOM_PERSPECTIVES.AfficheTiers (ARow: Longint): boolean;
Var Select : String;
    Q : TQuery;
begin
result:=True;
if GS.Cells[1,ARow] <> '' then
    begin
    Select :='SELECT T_LIBELLE,T_CODEPOSTAL,T_VILLE FROM TIERS WHERE T_TIERS="'+GS.Cells[1,ARow]+'" AND T_NATUREAUXI = "CON"';
    Q := OpenSQL(Select, True);
    if not Q.Eof then
        begin
        GS.Cells[2,ARow]:=Q.Fields[0].AsString;
        GS.Cells[3,ARow]:=Q.Fields[1].AsString;
        GS.Cells[4,ARow]:=Q.Fields[2].AsString;
        end
        else
        result:=false;
    Ferme(Q) ;
    end;
end;

procedure TOM_PERSPECTIVES.ContactNom;
var Q : TQuery;
begin
SetControlText ('LECONTACT','');
if (GetField('RPE_AUXILIAIRE') <> '') and (GetField('RPE_NUMEROCONTACT')<>0) then
   begin
   Q:=OpenSql('Select C_NOM from CONTACT where C_TYPECONTACT="T" AND C_AUXILIAIRE="'+GetField('RPE_AUXILIAIRE')+'" AND C_NUMEROCONTACT='+intToStr(GetField('RPE_NUMEROCONTACT')),TRUE);
   if not Q.EOF then
      begin
      SetControlText ('LECONTACT',Q.FindField('C_NOM').asstring);
      end;
   ferme(Q);
   end;
end;

procedure TOM_PERSPECTIVES.SetArguments(StSQL : string);
var Critere,ChampMul,ValMul : string ;
    x,y : integer ;
    Ctrl : TControl;
    Fiche : TFFiche;
begin
SetControlVisible('BSTOP',TRUE);
DS.Edit;
Fiche := TFFiche(ecran);
Repeat
    Critere:=Trim(ReadTokenPipe(StSQL,'|')) ;
    if Critere<>'' then
        begin
        x:=pos('=',Critere);
        if x<>0 then
           begin
           ChampMul:=copy(Critere,1,x-1);
           ValMul:=copy(Critere,x+1,length(Critere));
           y := pos(',',ValMul);
           if y<>0 then ValMul:=copy(ValMul,1,length(ValMul)-1);
           if copy(ValMul,1,1)='"' then ValMul:=copy(ValMul,2,length(ValMul));
           if copy(ValMul,length(ValMul),1)='"' then ValMul:=copy(ValMul,1,length(ValMul)-1);
           SetField(ChampMul,ValMul);
           Ctrl:=TControl(Fiche.FindComponent(ChampMul));
           if Ctrl=nil then exit;
{$IFDEF EAGLCLIENT}
           if (Ctrl is TCustomCheckBox) or (Ctrl is THValComboBox) Or (Ctrl is TCustomEdit) then
           begin
            TEdit(Ctrl).Font.Color:=clRed;
            SetControlText(ChampMul,ValMul);
           end
           else if Ctrl is TSpinEdit then TSpinEdit(Ctrl).Font.Color:=clRed
           else if (Ctrl is TCheckBox) or (Ctrl is THValComboBox) Or (Ctrl is THEdit)Or (Ctrl is THNumEdit)then
              begin
              TSpinEdit(Ctrl).Font.Color:=clRed;
              SetControlText(ChampMul,ValMul);
              end;
{$ELSE}
           if (Ctrl is TDBCheckBox) or (Ctrl is THDBValComboBox) Or (Ctrl is THDBEdit) then TEdit(Ctrl).Font.Color:=clRed
           else if Ctrl is THDBSpinEdit then THDBSpinEdit(Ctrl).Font.Color:=clRed
           else if (Ctrl is TCheckBox) or (Ctrl is THValComboBox) Or (Ctrl is THEdit)Or (Ctrl is THNumEdit)then
              begin
              THDBSpinEdit(Ctrl).Font.Color:=clRed;
              SetControlText(ChampMul,ValMul);
              end;
{$ENDIF}
           end;
        end;
until  Critere='';
end;
procedure TOM_PERSPECTIVES.AddRemoveItemFromPopup(stPopup,stItem : string; bVisible : boolean);
var pop : TPopupMenu ;
    i : integer;
    st : string;
begin
pop :=TPopupMenu(GetControl(stPopup) ) ;
if pop<> nil  then
   for i:=0 to pop.items.count-1 do
       begin
       st:=uppercase(pop.items[i].name);
       if st=stItem then  begin pop.items[i].visible:=bVisible; break; end;
       end;
end ;

function Perspective_MyAfterImport (Sender: TObject) : string;
var  OM : TOM ;
begin
result := '';
if sender is TFFICHE then OM := TFFICHE(Sender).OM else exit;
if (OM is TOM_PERSPECTIVES) then result := TOM_PERSPECTIVES(OM).GetInfoGed else exit;
end;

procedure Perspective_GestionBoutonGED (Sender: TObject);
var  OM : TOM ;
begin
if sender is TFFICHE then OM := TFFICHE(Sender).OM else exit;
if (OM is TOM_PERSPECTIVES) then TOM_PERSPECTIVES(OM).GestionBoutonGED else exit;
end;

function TOM_PERSPECTIVES.GetInfoGed : string;
begin
if not(DS.State in [dsInsert]) then result := 'Tiers='+GetField('RPE_TIERS')+';'+'Proposition='+intTostr(GetField('RPE_PERSPECTIVE'));
end;

procedure TOM_PERSPECTIVES.RTAppelGED_OnClick (Sender: TObject) ;
begin
if (ds<>nil) and not(DS.State in [dsInsert]) then
   begin
   AGLLanceFiche('RT','RTRECHDOCGED','RTD_TIERS='+GetField ('RPE_TIERS')+';RTD_PERSPECTIVE='+IntToStr(GetField ('RPE_PERSPECTIVE')),'','Objet=PRO;Tiers='+GetField ('RPE_TIERS'));
   GestionBoutonGED;
   end;
end;

Procedure TOM_PERSPECTIVES.GestionBoutonGED;
BEGIN
if (GetParamSocSecur('SO_RTGESTIONGED',False)) and (ds<>nil) and not(DS.State in [dsInsert]) and (ExisteSQL('SELECT RTD_DOCID FROM RTDOCUMENT WHERE RTD_TIERS="'+GetField('RPE_TIERS')+'" AND RTD_PERSPECTIVE='+IntToStr(GetField ('RPE_PERSPECTIVE')))) then
     SetControlVisible ('BDOCGEDEXIST',True)
else SetControlVisible ('BDOCGEDEXIST',False);
END;

{ GC/GRC : MNG / gestion des alertes }
procedure TOM_PERSPECTIVES.ListAlerte_OnClick_RPE(Sender: TObject);
begin
if (GetField('RPE_PERSPECTIVE') <> 0) and(AlerteActive(TableToPrefixe(TableName))) then
   AGLLanceFiche('Y','YALERTES_MUL','YAL_PREFIXE=RPE','','ACTION=CREATION;MONOFICHE;CHAMP=RPE_PERSPECTIVE;VALEUR='
      +IntToStr(GetField('RPE_PERSPECTIVE'))+';LIBELLE='+GetField('RPE_LIBELLE')) ;
end ;

procedure TOM_PERSPECTIVES.Alerte_OnClick_RPE(Sender: TObject);
begin
if (GetField('RPE_PERSPECTIVE') <> 0) and(AlerteActive(TableToPrefixe(TableName))) then
   AGLLanceFiche('Y','YALERTES','','','ACTION=CREATION;MONOFICHE;CHAMP=RPE_PERSPECTIVE;VALEUR='
   +IntToStr(GetField('RPE_PERSPECTIVE'))+';LIBELLE='+GetField('RPE_LIBELLE')) ;
VH_EntPgi.TobAlertes.ClearDetail;
end;

procedure TOM_PERSPECTIVES.RTPropoAppelParamCL(bCreation : boolean);
var TobChampsProFille : tob;
  StAction: string;
  { mng_fq012;10859 }
  bcreat:boolean;
begin
  { mng_fq012;10859 }
  bcreat:=false;
  if ds.state = dsInsert then bcreat:=true;
  VH_RT.TobChampsPro.Load;
  TobChampsProFille:=VH_RT.TobChampsPro.FindFirst(['CO_CODE'], ['V'], TRUE);
  if bCreation then
    begin
    if not Assigned(TobChampsProFille.FindFirst(['RDE_DESC','RDE_OBLIGATOIRE'],['V','X'],TRUE)) then exit;
    end
  else
    if (TobChampsProFille = Nil ) or (TobChampsProFille.detail.count = 0 ) then
        begin
        PGIInfo('Le paramétrage de cette saisie n''a pas été effectué','');
        exit;
        end
    else
    { mng_fq012;10858 }
      if not TFFiche(Ecran).Bouge(TNavigateBtn(nbPost)) then exit ;

  StAction:='';
  if (not self.ModifAutorisee) or (TFFiche(ecran).TypeAction=taConsult) then
     StAction:='ACTION=CONSULTATION;';
  { mng_fq012;10859 }
  if (not bcreat) or (not Assigned(TobChampsProFille.FindFirst(['RDE_DESC','RDE_OBLIGATOIRE'],['V','X'],TRUE))) then
    AglLancefiche('RT','RTPARAMCL','','',StAction+'FICHEPARAM='+TFFiche(Ecran).Name+';FICHEINFOS='+IntToStr(GetField('RPE_PERSPECTIVE'))+iif (bCreation,';EXISTOBLIG','')) ;
end;

procedure TOM_PERSPECTIVES.BIMPRIMEFICHE_OnClick(Sender: TObject);
var stArgument : string;
begin
{$IFDEF AFFAIRE}
  if ( (ctxAffaire in V_PGI.PGIContexte) or (ctxGCAFF in V_PGI.PGIContexte) ) and
     ( VH_GC.GASeria=True ) then stArgument := ''
  else stArgument := 'NOTAFF;';
{$ENDIF}
{$IFDEF GRCLIGHT}
if not GetParamsocSecur('SO_CRMACCOMPAGNEMENT',False) then
  stArgument := stArgument + 'NOTPROJET;NOTOPER;'
else
  if (Not GetParamSocSecur ('SO_RTPROJGESTION',False)) then stArgument := stArgument + 'NOTPROJET;';
{$ELSE}
if (Not GetParamSocSecur ('SO_RTPROJGESTION',False)) then stArgument := stArgument + 'NOTPROJET;';
{$ENDIF}
AGLLanceFiche ('RT','RTPARAMEDITPROPO','RPE_PERSPECTIVE='+IntToStr(GetField('RPE_PERSPECTIVE')),'',stArgument)
end;


//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
// Fonctions pour Script

procedure AGLPiecePerspective( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
    // fct qui publie la fct de ctrl modification ressource
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_PERSPECTIVES)
    then TOM_PERSPECTIVES(OM).PiecePerspective(Parms[1])
    else exit;
end;

procedure AGLVariantePerspective( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_PERSPECTIVES)
    then TOM_PERSPECTIVES(OM).VariantePerspective(Parms[1])
    else exit;
end;

procedure AGLEnvoiMessage( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_PERSPECTIVES) then TOM_PERSPECTIVES(OM).EnvoiMessage else exit;
end;

procedure AGLChargeGrid( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_PERSPECTIVES) then TOM_PERSPECTIVES(OM).ChargeGrid(Parms[1]) else exit;
end;

procedure AGLClickPrincipal (parms:array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_PERSPECTIVES) then TOM_PERSPECTIVES(OM).ClickPrincipal else exit;
end;

procedure AGLRTnouvellePiece( parms: array of variant; nb: integer )  ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_PERSPECTIVES) then TOM_PERSPECTIVES(OM).NouvellePiece;
end;

procedure AGLRTnouvelleAffaire( parms: array of variant; nb: integer )  ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_PERSPECTIVES) then TOM_PERSPECTIVES(OM).NouvelleAffaire;
end;

procedure AGLRTPerspectiveGenereAction( parms: array of variant; nb: integer )  ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_PERSPECTIVES) then TOM_PERSPECTIVES(OM).PerspGenereAction;
end;

procedure AGLContactNom( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_PERSPECTIVES) then TOM_PERSPECTIVES(OM).ContactNom else exit;
end;

procedure AGLRTPropoAppelParamCL( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_PERSPECTIVES) then TOM_PERSPECTIVES(OM).RTPropoAppelParamCL(false) else exit;
end;


// *****************************************************************************
// ********************** gestion Isoflex **************************************
// *****************************************************************************

procedure TOM_PERSPECTIVES.GereIsoflex;
var bIso: Boolean;
  MenuIso: TMenuItem;
begin
  MenuIso := TMenuItem(GetControl('mnSGED'));
  bIso := AglIsoflexPresent;
  if MenuIso <> nil then MenuIso.Visible := bIso;
end;

procedure Rpe_AppelIsoFlex(parms: array of variant; nb: integer);
var F: TForm;
begin
  F := TForm(Longint(Parms[0]));
  //GP_20071226_DS_GP14537
  if not pos('RTPERSPECTIVES', UpperCase(F.Name)) = 1 then exit;
  AglIsoflexViewDoc(NomHalley, F.Name, 'PERSPECTIVES', 'RPE_CLE1', 'RPE_PERSPECTIVE', string(Parms[1]), '');
end;

Initialization
registerclasses([TOM_PERSPECTIVES]) ;
RegisterAglProc( 'PiecePerspective', TRUE , 1, AGLPiecePerspective);
RegisterAglProc( 'RTVariantePerspective', TRUE , 1, AGLVariantePerspective);
RegisterAglProc( 'EnvoiMessage', TRUE , 0, AGLEnvoiMessage);
RegisterAglProc( 'ClickPrincipal', TRUE , 0, AGLClickPrincipal);
RegisterAglProc( 'ChargeGrid', TRUE , 1, AGLChargeGrid);
RegisterAglProc( 'RTNouvellePiece', True,0,AGLRTNouvellePiece) ;
RegisterAglProc( 'RTNouvelleAffaire', True,0,AGLRTNouvelleAffaire) ;
RegisterAglProc( 'RTPerspectiveGenereAction', True,0,AGLRTPerspectiveGenereAction) ;
RegisterAglProc( 'ContactNom', TRUE , 0, AGLContactNom);
RegisterAglProc('Rpe_AppelIsoFlex', TRUE, 1, Rpe_AppelIsoFlex);
RegisterAglProc( 'RTPropoAppelParamCL', True,0,AGLRTPropoAppelParamCL) ;
end.
