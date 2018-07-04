unit UtomProjet;

interface

uses Classes,forms,sysutils,AglIsoflex,
{$IFDEF EAGLCLIENT}
     eFiche,MaineAGL, UtileAGL,
{$ELSE}
     db,hdb,{$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}Fiche,FE_Main,DBCtrls,
{$ENDIF}
     ComCtrls,HMsgBox,  menus,
     HCtrls,HEnt1,UTOM,
     UTob,M3FP,Grids,graphics,
     Windows  ;

Type
    TOM_PROJETS = Class (TOM)
    Public
    stTiers : String;
    procedure OnUpdateRecord ; override ;
    Procedure OnDeleteRecord ; override ;
    procedure OnChangeField (F : TField) ; override ;
    procedure OnArgument (Arguments : String ); override ;
    procedure OnLoadRecord ; override ;
    procedure OnLoadAlerte  ; override ;
    procedure OnNewRecord ; override ;
    private
    LesColonnes: string ;
    GS : THGRID ;
    NochangeProspect,NochangeContact : Boolean;
    iNumeroContact : Integer;
    stAction : String;
    procedure GSLigneDClick (Sender: TObject);
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure DessineCell (ACol,ARow : Longint; Canvas : TCanvas;
                                     AState: TGridDrawState);
    procedure RTAffichagePropositions;
    procedure ProjetNomContact;
		procedure RPJ_TIERS_OnElipsisClick(Sender:TObject);
        // Gestion Isoflex
    procedure GereIsoflex;
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure RTProjetAppelParamCL;
  {$IFNDEF EAGLSERVER}
    function TestAlerteProjet (CodeA : String) : boolean;
    procedure ListAlerte_OnClick_RPJ(Sender: TObject);
    procedure Alerte_OnClick_RPJ(Sender: TObject);
  {$ENDIF EAGLSERVER}

END ;

procedure AGLRTAffichagePropositions(parms:array of variant; nb: integer ) ;

const
	// libellés des messages
	TexteMessage: array[1..9] of string 	= (
          {1}        'Vous devez renseigner la date de début'
          {2}        ,'Vous devez renseigner le tiers'
          {3}        ,'Vous ne pouvez pas effacer ce projet, il est actuellement utilisé dans les actions.'
          {4}        ,'Vous ne pouvez pas effacer ce projet, il est actuellement utilisé dans les propositions.'
          {5}        ,'Le code tiers n''existe pas.'
          {6}        ,'Le code tiers est obligatoire.'
          {7}        ,'Opération impossible.Ce tiers est fermé'
          {8}        ,'Le responsable n''existe pas'
          {9}        ,'Vous n''êtes pas autorisé à saisir ce projet'
          );

var
  Cell_Perspective  : integer;

implementation
uses
  SaisUtil,TiersUtil,ParamSoc,EntRT,UtilGC,UtilRT, UtilConfid,AGLInitGC
{$IFNDEF EAGLSERVER}
  ,UtilAlertes,YAlertesConst,EntPgi
{$ENDIF EAGLSERVER}
;
procedure AGLRTAffichagePropositions( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_PROJETS) then TOM_PROJETS(OM).RTAffichagePropositions else exit;
end;

procedure AGLProjetNomContact( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_PROJETS) then TOM_PROJETS(OM).ProjetNomContact else exit;
end;

procedure TOM_PROJETS.OnArgument (Arguments : String );
var x : integer;
    Critere,ChampMul,ValMul : string;
begin
inherited ;
    stTiers := '';
    iNumeroContact:=0;
    stAction := '';
    Repeat
        Critere:=uppercase(ReadTokenSt(Arguments)) ;
        if Critere<>'' then
            begin
            x:=pos('=',Critere);
            if x<>0 then
               begin
               ChampMul:=copy(Critere,1,x-1);
               ValMul:=copy(Critere,x+1,length(Critere));
               if ChampMul='RPJ_TIERS' then StTiers := ValMul;
               if ChampMul='RPJ_NUMEROCONTACT' then iNumeroContact := strToInt(ValMul);
               if ChampMul='ACTION' then stAction := ValMul;
               end;
            x:=pos('NOCHANGEPROSPECT',Critere);
            if x<>0 then
               NoChangeProspect := True ;
            x:=pos('FICHECONTACT',Critere);
            if x<>0 then
               NoChangeContact := True ;
            end;
    until  Critere='';

    if (stAction = 'CONSULTATION') and ( not VH_RT.RTCreatPropositions )
        then SetControlVisible ('BPROPOSITION',false);
    SetControlText ('TIERS','');
//    SetControlText ('NUMEROCONTACT','0');
    if stTiers <> '' then SetControlText ('TIERS',stTiers);
    if iNumeroContact <> 0 then SetControlText ('NUMEROCONTACT',intToStr(iNumeroContact));


    LesColonnes :='FIXED;PRINCIPALE;RPE_PERSPECTIVE;T_LIBELLE;RPE_TYPEPERSPECTIV;RPE_LIBELLE;RPE_DATEREALISE;RPE_MONTANTPER;RPE_ETATPER;RPE_BLOCNOTE;';
    GS:=THGRID(GetControl('GR_LISTEPROPOSITIONS'));
    GS.OnRowEnter:=GSRowEnter ;
    GS.OnRowExit:=GSRowExit ;
    GS.OnDblClick:=GSLigneDClick ;
    GS.PostDrawCell:= DessineCell;
    GS.ColWidths[0]:=15;
    GS.ColWidths[3]:=100;
    GS.ColWidths[4]:=100;
    GS.ColWidths[5]:=160;
    GS.ColWidths[6]:=85;
    GS.ColWidths[7]:=100;
    GS.ColWidths[8]:=90;
    GS.ColWidths[9]:=50;
//    GS.ColAligns[8]:=taCenter;
    GS.ColWidths[2]:=70; Cell_perspective:=2;
    GS.ColFormats[2]:='0';
    GS.ColFormats[4]:='CB=RTTYPEPERSPECTIVE';
    GS.ColFormats[8]:='CB=RTETATPERSPECTIVE';
    GS.ColAligns[7]:=taRightJustify;
    GS.ColFormats[7]:='#,##0';
    GS.ColWidths[1]:=20;
    GS.ColTypes[1] := 'B';
    GS.ColAligns[1]:=taCenter;
    GS.ColFormats[1]:=intToStr(integer(csCoche));
    AffecteGrid(GS,taConsult) ;
    TFFiche(Ecran).Hmtrad.ResizeGridColumns(GS) ;
    GereIsoflex;
{$ifdef GIGI}
 SetControlText ('TRPJ_NUMEROCONTACT',TraduireMemoire('Contact'));
 If (Not GetParamSocSecur ('SO_AFRTPROPOS',False)) then
   begin
    SetControlVisible('BPROPOSITION',false);
    SetControlVisible('PPROPOSITION',false);
   end;
{$ENDIF}

  if ecran is TFFiche then
    TFFiche(Ecran).OnKeyDown:=FormKeyDown ;

// Pl le 19/05/07 : gestion des champs libres seulement pour KPMG pour l'instant
//  if (GetParamSocSecur ('SO_AFCLIENT', 0) = 8 ) then
    GCMAJChampLibre(TForm(Ecran), False, 'COMBO', 'RPJ_RPJTABLELIBRE', 5, '_');
//  else
//    SetControlVisible('PTABLESLIBRES', false);
{$IFNDEF EAGLSERVER}
  if Assigned(GetControl('MnAlerte')) then
    if AlerteActive('RPJ') then
      TMenuItem(GetControl('MnAlerte')).OnClick := Alerte_OnClick_RPJ
    else
      TMenuItem(GetControl('MnAlerte')).visible:=false;

  if Assigned(GetControl('MnListAlerte')) then
    if AlerteActive('RPJ') then
      TMenuItem(GetControl('MnListAlerte')).OnClick := ListAlerte_OnClick_RPJ
    else
      TMenuItem(GetControl('MnListAlerte')).visible:=false;

  if Assigned(GetControl('MnGestAlerte')) and Assigned(GetControl('MnAlerte'))
     and Assigned(GetControl('MnListAlerte')) then
         TMenuItem(GetControl('MnGestAlerte')).visible := (TMenuItem(GetControl('MnAlerte')).visible)
          and (TMenuItem(GetControl('MnListAlerte')).visible);
{$ENDIF EAGLSERVER}
{$IFDEF EAGLCLIENT}
  if Assigned(GetControl('RPJ_TIERS')) then
    THEdit(GetControl('RPJ_TIERS')).OnElipsisClick:= RPJ_Tiers_OnElipsisClick;
{$ELSE}
  if Assigned(GetControl('RPJ_TIERS')) then
    THDBEdit(GetControl('RPJ_TIERS')).OnElipsisClick:= RPJ_Tiers_OnElipsisClick;
{$ENDIF}

end;

procedure TOM_PROJETS.OnNewRecord;
begin
inherited;
SetField('RPJ_DATEDEBUT', V_PGI.DateEntree);
SetField('RPJ_TIERS',stTiers);
if (stTiers <> '') then
   begin
   SetField('RPJ_AUXILIAIRE',TiersAuxiliaire (GetField('RPJ_TIERS'), False));
   end;
SetControlText ('LECONTACT','');
if iNumeroContact <> 0 then SetField('RPJ_NUMEROCONTACT',iNumeroContact);
if GetParamsocSecur('SO_RTPROJRESP','COM') = 'UTI' then SetField('RPJ_INTERVENANT',VH_RT.RTResponsable);
SetControlEnabled('RPJ_TIERS',True) ;
GS.videpile (False);
end;

procedure TOM_PROJETS.OnUpdateRecord;
var Q : TQuery;
    Select,CodeProjet, NumChrono : String;
    ExisteOblig : boolean;
begin
inherited;
if GetField ('RPJ_DATEDEBUT') = iDate1900 then
   begin
   Lasterror:=1;
   LastErrorMsg:=TexteMessage[LastError];
   SetFocusControl('RPJ_DATEDEBUT') ;
   exit;
   end;
if (GetField('RPJ_TIERS')<>'') then
   begin
   if not ExisteSQL ('SELECT T_LIBELLE FROM TIERS WHERE T_AUXILIAIRE="'+GetField('RPJ_AUXILIAIRE')+'"') then
     begin
     LastError := 5;
     LastErrorMsg:=TexteMessage[LastError];
     SetFocusControl('RPJ_TIERS');
     exit;
     end;
   end
   else
   if GetParamSocSecur('SO_RTPROJATTACHTIERS',False) then
     begin
       LastError := 6;
       LastErrorMsg:=TexteMessage[LastError];
       SetFocusControl('RPJ_TIERS');
       exit;
     end;

if ((DS.State in [dsInsert]) and (GetField('RPJ_TIERS')<>'')) then
   begin
   if not RTDroitModifTiers(GetField('RPJ_TIERS')) then
      begin
      SetFocusControl('RPJ_TIERS');
      LastError :=9;
      PGIBox(TexteMessage[LastError],'Accès en création');
      exit;
      end;
   end;

if (GetField('RPJ_TIERS')<>'') then
   begin
   Select :='SELECT T_FERME FROM TIERS WHERE T_AUXILIAIRE="'+GetField('RPJ_AUXILIAIRE')+'"';
   Q := OpenSQL(Select, True);
   if (Q.Eof) or (Q.FindField('T_FERME').asstring = 'X') then
      begin
      Ferme(Q) ;
      SetFocusControl('RPJ_TIERS');
      LastError := 7;
      LastErrorMsg:=TexteMessage[LastError];
      exit;
      end;
   Ferme(Q) ;
   end;
if ((GetField ('RPJ_INTERVENANT') <> '') and  (Not ExisteSQL ('SELECT ARS_LIBELLE FROM RESSOURCE WHERE ARS_RESSOURCE="'+GetField('RPJ_INTERVENANT')+'"'))) then
   begin
   SetFocusControl('RPJ_INTERVENANT');
   LastError := 8;
   LastErrorMsg:=TexteMessage[LastError];
   exit;
   end;
{$IFNDEF EAGLSERVER}
  if  (Assigned(Ecran)) and (not V_Pgi.SilentMode) and (AlerteActive(TableToPrefixe(TableName))) then
       if (not Inserting) then
         begin
         if not TestAlerteProjet (CodeModification+';'+CodeModifChamps) then
           begin
           LastError:=99;
           exit;
           end;
         end
       else
         if not TestAlerteProjet (CodeCreation) then
           begin
           LastError:=99;
           exit;
           end;
{$ENDIF !EAGLSERVER}

if GetParamsocSecur('SO_RTNUMPROJETAUTO',False) = TRUE then
  begin
  if (ds<>nil) and (ds.state = dsinsert) and (GetField ('RPJ_PROJET')= '') then
     begin
     CodeProjet:=AttribNewCode('PROJETS','RPJ_PROJET',0,'',GetParamsocSecur('SO_RTCOMPTEURPROJET',''),'');
     SetField ('RPJ_PROJET', CodeProjet);
{$IFDEF EAGLCLIENT}
     SetControlText('RPJ_PROJET',CodeProjet);
{$ENDIF}
     NumChrono:=ExtraitChronoCode(CodeProjet);
     SetParamSoc('SO_RTCOMPTEURPROJET', NumChrono) ;
     end;
  end;

  if (DS<>nil) and (DS.State=dsInsert) and (GetParamSocSecur('SO_RTGESTINFOS00Q',False)) then
    begin
    ExisteOblig:=ExisteSQL('SELECT RDE_OBLIGATOIRE FROM RTINFOSDESC WHERE RDE_DESC="Q" AND RDE_OBLIGATOIRE="X" ');
    if ExisteOblig then
       AglLancefiche('RT','RTPARAMCL','','','FICHEPARAM='+TFFiche(Ecran).Name+';FICHEINFOS='+GetField ('RPJ_PROJET')+';EXISTOBLIG') ;
    end;

end;

procedure TOM_PROJETS.OnDeleteRecord  ;
begin
Inherited ;

if ExisteSQL('SELECT RAC_LIBELLE FROM ACTIONS WHERE RAC_PROJET="'
     +GetField('RPJ_PROJET')+'" ') then
   BEGIN
   LastError:=3;
   LastErrorMsg:=TexteMessage[LastError];
   exit ;
   end ;
if ExisteSQL('SELECT RPE_LIBELLE FROM PERSPECTIVES WHERE RPE_PROJET="'
     +GetField('RPJ_PROJET')+'" ') then
   BEGIN
   LastError:=4;
   LastErrorMsg:=TexteMessage[LastError];
   exit ;
   end ;

  if (not V_Pgi.SilentMode) and (not AfterInserting)
    and (AlerteActive (TableToPrefixe(TableName))) then
      if (not TestAlerteProjet(CodeSuppression)) then
        begin
        LastError := 99;
        exit;
        end;

if GetParamSocSecur('SO_RTGESTINFOS00Q',False) then
   ExecuteSQL('DELETE FROM RTINFOS00Q where RDQ_CLEDATA="'+GetField('RPJ_PROJET')+'"') ;
end;

procedure TOM_PROJETS.OnChangeField(F: TField);
var QQ : tquery;
    Select : string;
begin
inherited;
//if (DS.State = dsInsert) then
//  begin
   if (F.FieldName = 'RPJ_TIERS') then
      begin
      if DS.State in [dsInsert,dsEdit] then
        begin
        if (GetField('RPJ_TIERS')<>'') then
           SetField('RPJ_AUXILIAIRE',TiersAuxiliaire (GetField('RPJ_TIERS'), False))
        else
           SetField('RPJ_AUXILIAIRE','');
        end;
        if (DS.State = dsInsert) then
          begin
              //mcd 30/11/2005 il faut prendre en compte le paramétrage possible intervenent que ressource
          if (( GetParamsocSecur('SO_RTPROJRESP','COM') = 'RE1' )
           or ( GetParamsocSecur('SO_RTPROJRESP','COM') = 'RE2' )
           or ( GetParamsocSecur('SO_RTPROJRESP','COM') = 'RE3' ))
           and (CtxAffaire in V_PGI.PGIContexte)
           and (GetField('RPJ_TIERS')<>'')
           and (GetControlText('RPJ_INTERVENANT')='') then
           begin
             Select :='SELECT YTC_RESSOURCE1,YTC_RESSOURCE2,YTC_RESSOURCE3 FROM  TIERSCOMPL WHERE YTC_TIERS="'+GetField('RPJ_TIERS')+'"';
             QQ := OpenSQL(Select, True);
             if not QQ.Eof then
                  begin
                  if ( GetParamsocSecur('SO_RTPROJRESP','COM') = 'RE1' ) then SetField('RPJ_INTERVENANT',QQ.FindField('YTC_RESSOURCE1').asstring);
                  if ( GetParamsocSecur('SO_RTPROJRESP','COM') = 'RE2' ) then SetField('RPJ_INTERVENANT',QQ.FindField('YTC_RESSOURCE2').asstring);
                  if ( GetParamsocSecur('SO_RTPROJRESP','COM') = 'RE3' ) then SetField('RPJ_INTERVENANT',QQ.FindField('YTC_RESSOURCE3').asstring);
                  end;
             Ferme(QQ);
             end;
                //fin mcd 30/11/2005
          if (GetParamsocSecur('SO_RTPROJRESP','COM') = 'COM') and
             (GetField('RPJ_TIERS')<>'') then
             begin
             SetField('RPJ_INTERVENANT', RTRechResponsable (GetField('RPJ_TIERS')));
             end;
         end;
      end;
//   end;
end;
procedure TOM_PROJETS.OnLoadAlerte;
begin
  {$IFNDEF EAGLSERVER}
  if (not V_Pgi.SilentMode) and (AlerteActive(TableToPrefixe(TableName))) and (not AfterInserting )  then
     TestAlerteProjet(CodeOuverture+';'+CodeDateAnniv);
  {$ENDIF !EAGLSERVER}
end;

procedure TOM_PROJETS.OnLoadRecord;
var F : TForm;
begin
inherited;
if  GetField('RPJ_PROJET') <> '' then
    RTAffichagePropositions;
F := TForm (Ecran);
if (TFFICHE(F).TypeAction = TaModif) and (not (DS.State in [dsInsert])) then
    TFFiche(Ecran).Pages.ActivePage:=TTabSheet(Ecran.FindComponent('PPROPOSITION'))
     else
    TFFiche(Ecran).Pages.ActivePage:=TTabSheet(Ecran.FindComponent('PGeneral'));
if ((not(DS.State in [dsInsert])) and ((GetField('RPJ_TIERS') <> '') or ((GetField('RPJ_TIERS') = '') and (GetParamSocSecur ('SO_RTPROJMULTITIERS',True))))) or (NochangeProspect = True) then
   begin
   SetControlEnabled('RPJ_TIERS',False) ;
   end;
if (GetParamsocSecur('SO_RTNUMPROJETAUTO',False)  = TRUE) then
   begin
    SetControlEnabled('RPJ_PROJET', FALSE);
    if ds.state in [dsinsert] then SetFocusControl('RPJ_LIBELLE');
   end;
if NoChangeContact = True then
   SetControlEnabled('LECONTACT',False) ;
ProjetNomContact();
AppliquerConfidentialite(Ecran,'');
end;

procedure TOM_PROJETS.RTAffichagePropositions;
var Q : TQuery;
    Select : String;
    TOBProp : TOB;
    i : Integer;
    Principale : String;
begin
inherited ;
TOBProp:=tob.create('Liste des propositions',Nil,-1) ;
Select := 'SELECT RPE_PERSPECTIVE,RPE_LIBELLE,RPE_INTERVENANT,RPE_TYPEPERSPECTIV,RPE_ETATPER,RPE_MONTANTPER,RPE_BLOCNOTE,RPE_DATEREALISE,'+
       'RPE_VARIANTE,T_LIBELLE FROM PERSPECTIVES LEFT JOIN TIERS on T_AUXILIAIRE=RPE_AUXILIAIRE '+
       'WHERE RPE_PROJET="'+GetField('RPJ_PROJET')+'"';
if GetControlText('PROPOPRINCIPALE') = 'X' then Select := Select + ' AND (RPE_VARIANTE=0 or RPE_PERSPECTIVE=RPE_VARIANTE)';
if StTiers <> '' then
   begin
   Select := Select + ' AND RPE_TIERS ="'+stTiers+'"';
   if iNumeroContact <> 0 then Select := Select + ' AND RPE_NUMEROCONTACT ='+IntToStr(iNumeroContact) ;
   end;
Select := Select + ' ORDER BY RPE_PERSPECTIVE DESC';

Q:=OpenSQL(Select, True);

TOBProp.LoadDetailDB('','','',Q,false,true) ;
Ferme(Q);

if (TOBProp.Detail.count > 0) then
   begin
   TOBProp.Detail[0].AddChampSup('PRINCIPALE', True);
   for i:=0 to TOBProp.Detail.count-1 do
       begin
       Principale := '-';
       if (TOBProp.detail[i].GetValue('RPE_VARIANTE')= 0) or (TOBProp.detail[i].GetValue('RPE_VARIANTE') = TOBProp.detail[i].GetValue('RPE_PERSPECTIVE')) then
           begin
           Principale := 'X';
           end;
       TOBProp.Detail[i].PutValue('PRINCIPALE', Principale);
       end;
   end;

TOBProp.PutGridDetail(GS,False,False,LesColonnes,True);

SetControlText('MONTANTGLOBAL',FloatToStrF(TobProp.Somme('RPE_MONTANTPER',['PRINCIPALE'],['X'],TRUE,FALSE),ffNumber,15,0));
TOBProp.free;
end;

procedure TOM_PROJETS.ProjetNomContact;
var Q : TQuery;
begin
SetControlText ('LECONTACT','');
if (GetField('RPJ_AUXILIAIRE') <> '') and (GetField('RPJ_NUMEROCONTACT')<>0) then
   begin
   Q:=OpenSql('Select C_NOM from CONTACT where C_TYPECONTACT="T" AND C_AUXILIAIRE="'+GetField('RPJ_AUXILIAIRE')+'" AND C_NUMEROCONTACT='+intToStr(GetField('RPJ_NUMEROCONTACT')),TRUE);
   if not Q.EOF then
      begin
      SetControlText ('LECONTACT',Q.FindField('C_NOM').asstring);
      end;
   ferme(Q);
   end;
end;

procedure TOM_PROJETS.RPJ_TIERS_OnElipsisClick(Sender: TObject);
begin
  if not(DS.State in [dsInsert,dsEdit]) then
    DS.Edit;
	DispatchRecherche(THCritMaskEdit(getControl('RPJ_TIERS')), 2, '', '', '');
end;

procedure TOM_PROJETS.GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
GS.InvalidateRow(ou) ;
end;

procedure TOM_PROJETS.GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
GS.InvalidateRow(ou) ;
end;

procedure TOM_PROJETS.GSLigneDClick (Sender: TObject);
var St_Persp,stArg : string;
begin
  St_Persp := string(GS.CellValues[Cell_Perspective, GS.Row]);
  if (St_Persp = '') then Exit;
  stArg := 'ACTION=MODIFICATION';
  if stAction = 'CONSULTATION' then stArg := 'ACTION=CONSULTATION';
  AGLLanceFiche('RT','RTPERSPECTIVES','',St_Persp,stArg+';MONOFICHE');
  OnloadRecord;
end;

procedure TOM_PROJETS.DessineCell(ACol, ARow: Longint; Canvas : TCanvas; AState: TGridDrawState);
var Triangle : array[0..2] of TPoint ;
  Arect: Trect ;
begin
If Arow < GS.Fixedrows then exit ;
if (gdFixed in AState) and (ACol = 0) then
    begin
    Arect:=GS.CellRect(Acol,Arow) ;
    Canvas.Brush.Color := GS.FixedColor;
    Canvas.FillRect(ARect);
      if (ARow = GS.row) then
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
procedure TOM_PROJETS.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  Case Key of
      VK_F6 : {Infos compl.} if (ssAlt in Shift) then
         if GetParamSocSecur('SO_RTGESTINFOS00Q',False) then RTProjetAppelParamCL ;
{$IFNDEF EAGLSERVER}
    81 : {Ctrl + Q - Création d'1 alerte} if (ssCtrl in Shift) then
          begin
          Key:=0 ;
          Alerte_OnClick_RPJ(Sender);
          end;
    85 : {Ctrl + U - liste des alertes du tiers} if (ssCtrl in Shift) then
          begin
          Key:=0 ;
          ListAlerte_OnClick_RPJ(Sender);
          end;
{$ENDIF EAGLSERVER}
      END;
  if (ecran <> nil) then
    if ecran is TFFiche then
       TFFiche(ecran).FormKeyDown(Sender,Key,Shift);
end;

procedure TOM_PROJETS.RTProjetAppelParamCL;
var TobChampsProFille : tob;
  { mng_fq012;10859 }
  bcreat:boolean;
begin
//if VH_RT.TobChampsPro.detail[2].Detail.Count = 0 then
  VH_RT.TobChampsPro.Load;

  TobChampsProFille:=VH_RT.TobChampsPro.FindFirst(['CO_CODE'], ['Q'], TRUE);
  if (TobChampsProFille = Nil ) or (TobChampsProFille.detail.count = 0 ) then
      begin
      PGIInfo('Le paramétrage de cette saisie n''a pas été effectué','');
      exit;
      end;
  { mng_fq012;10859 }
  bcreat:=false;
  if ds.state = dsInsert then bcreat:=true;
  { mng_fq012;10858 }
  if not TFFiche(Ecran).Bouge(TNavigateBtn(nbPost)) then exit ;
  { mng_fq012;10859 }
  if (not bcreat) or (not Assigned(TobChampsProFille.FindFirst(['RDE_DESC','RDE_OBLIGATOIRE'],['Q','X'],TRUE))) then
    AglLancefiche('RT','RTPARAMCL','','','FICHEPARAM='+TFFiche(Ecran).Name+';FICHEINFOS='+GetField('RPJ_PROJET')) ;
end;

{$IFNDEF EAGLSERVER}
function TOM_PROJETS.TestAlerteProjet (CodeA : String) : boolean;
var TOBInfosCompl : tob;
    i : integer;
    F: TForm;
begin
  result:=true;
  F:=TForm(Ecran);
  { cas comme la duplication ou l'on passe dans le loadalerte alors que les champs ne sont pas renseignés }
  if (GetField('RPJ_PROJET') = '' ) then exit;
  if not (F is TfFiche) then exit;
  if assigned( TfFiche(F).TobFinale) then TfFiche(F).TobFinale.free;
  TfFiche(F).TobFinale:=TOB.create ('PROJETS',NIL,-1);

  TfFiche(F).TobFinale.GetEcran (TFfiche(Ecran),Nil);

  { si passage du load, on sauvegarde les tobs initiales }
  if pos(CodeOuverture,CodeA) > 0 then
    begin
    if assigned( TfFiche(F).TobOrigine) then TfFiche(F).TobOrigine.free;
    TfFiche(F).TobOrigine:=TOB.create ('PROJETS',NIL,-1);
    TfFiche(F).TobOrigine.GetEcran (F,Nil,true);
    end;

  if GetParamSocSecur('SO_RTGESTINFOS00Q',false) then
    begin
    TOBInfosCompl:= TOB.Create('RTINFOS00Q', nil, -1);
    if CodeA<>CodeCreation then
      TOBInfosCompl.selectDB ('"'+GetField('RPJ_PROJET')+'"',nil);
    for i := 1 to Pred(TOBInfosCompl.NbChamps) do
      TfFiche(F).TobFinale.AddChampSupValeur(TOBInfosCompl.GetNomChamp(i), TOBInfosCompl.GetValue(TOBInfosCompl.GetNomChamp(i)));
    if (pos(CodeOuverture,CodeA) > 0) then
      begin
      if (pos(CodeOuverture,CodeA) > 0) and assigned( TfFiche(F).TobOrigine) then
        for i := 1 to Pred(TOBInfosCompl.NbChamps) do
          TfFiche(F).TobOrigine.AddChampSupValeur(TOBInfosCompl.GetNomChamp(i), TOBInfosCompl.GetValue(TOBInfosCompl.GetNomChamp(i)));
      end;
    TOBInfosCompl.free;
    end;
  if pos(CodeOuverture,CodeA) > 0 then
    result:=ExecuteAlerteLoad(F,false)
  else
    if pos(CodeSuppression,CodeA) > 0 then
      result:=ExecuteAlerteDelete(F,false)
    else
      result:=ExecuteAlerteUpdate(F,false);
end;

{ GC/GRC : MNG / gestion des alertes }
procedure TOM_PROJETS.ListAlerte_OnClick_RPJ(Sender: TObject);
begin
if (GetField('RPJ_PROJET') <> '') and(AlerteActive(TableToPrefixe(TableName))) then
   AGLLanceFiche('Y','YALERTES_MUL','YAL_PREFIXE=RPJ','','ACTION=CREATION;MONOFICHE;CHAMP=RPJ_PROJET;VALEUR='
      +GetField('RPJ_PROJET')+';LIBELLE='+GetField('RPJ_LIBELLE')) ;
end ;

procedure TOM_PROJETS.Alerte_OnClick_RPJ(Sender: TObject);
begin
  if (GetField('RPJ_PROJET') <> '') and(AlerteActive(TableToPrefixe(TableName))) then
    AGLLanceFiche('Y','YALERTES','','','ACTION=CREATION;MONOFICHE;CHAMP=RPJ_PROJET;VALEUR='
    +GetField('RPJ_PROJET')+';LIBELLE='+GetField('RPJ_LIBELLE')) ;
  VH_EntPgi.TobAlertes.ClearDetail;
end;
{$ENDIF !EAGLSERVER}

procedure AGLRTProjetAppelParamCL( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_PROJETS) then TOM_PROJETS(OM).RTProjetAppelParamCL else exit;
end;
// *****************************************************************************
// ********************** gestion Isoflex **************************************
// *****************************************************************************

procedure TOM_PROJETS.GereIsoflex;
var bIso: Boolean;
  MenuIso: TMenuItem;
begin
  MenuIso := TMenuItem(GetControl('mnSGED'));
  bIso := AglIsoflexPresent;
  if MenuIso <> nil then MenuIso.Visible := bIso;
end;

procedure Rpj_AppelIsoFlex(parms: array of variant; nb: integer);
var F: TForm;
  Cle1: string;
begin
  F := TForm(Longint(Parms[0]));
  if (F.Name <> 'RTPROJETS') then exit;
  Cle1 := string(Parms[1]);
  AglIsoflexViewDoc(NomHalley, F.Name, 'PROJETS', 'RPJ_CLE1', 'RPJ_PROJET', Cle1, '');
end;
Initialization
registerclasses([TOM_PROJETS]) ;
RegisterAglProc( 'RTAffichagePropositions', TRUE , 0, AGLRTAffichagePropositions);
RegisterAglProc( 'ProjetNomContact', TRUE , 0, AGLProjetNomContact);
RegisterAglProc( 'RTProjetAppelParamCL', True,0,AGLRTProjetAppelParamCL) ;
RegisterAglProc('Rpj_AppelIsoFlex', TRUE, 1, Rpj_AppelIsoFlex);
end.

