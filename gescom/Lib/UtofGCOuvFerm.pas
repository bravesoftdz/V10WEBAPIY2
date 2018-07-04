unit UtofGCOuvFerm;

interface

uses  Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HMsgBox,UTOF,HQry,HEnt1,
{$IFDEF EAGLCLIENT}
      eMul,MaineAGL,
{$ELSE}
      mul,Fe_Main,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
      Hstatus,HDB,M3FP,UTOB,
{$IFDEF GRC}
      UtilSelection,
      UtilRT,
{$ENDIF}
      UtilPGI,TiersUtil,ParamSoc;

Type
    tof_gcouvferm = Class (TOF)
    private
           bFermer : boolean  ;
           stWhere,CodeTiers : string ;
    Public
        procedure OnLoad ; override ;
        procedure OnArgument (Arguments : String ); override ;
        procedure OuvreFerme;
        procedure SetAllOuvreFerme  ;
        procedure SetOuvreFerme  ;
     END;

implementation


procedure tof_gcouvferm.OnArgument (Arguments : String );
var    F : TForm;
    NomForme : string;
begin
inherited ;
F := TForm (Ecran);
{$IFDEF GRC}
if (ctxGRC in V_PGI.PGIContexte) and (pos('FOU', Arguments) = 0) then
   MulCreerPagesCL(F,'NOMFIC=GCTIERS');
if (ctxGRC in V_PGI.PGIContexte) and (pos('FOU', Arguments) <> 0)
   and (GetParamSoc('SO_RTGESTINFOS003') = True) then
    begin
    NomForme :=F.Name;
    F.Name := 'GCFOURNISSEURS';
    MulCreerPagesCL(F,'NOMFIC=GCFOURNISSEURS');
    F.Name:=NomForme;
    TFMul(F).Q.Liste:='RFMULOUVFERMFOURN'
    end;
{$ENDIF}
if pos('OUVRE',Arguments)>0 then
   begin
   bFermer:=False  ;
   SetControlText('T_FERME','X');
   if pos('GRC',Arguments)>0 then
       Ecran.Caption:='Activation des clients/prospects'
   else
   begin
      SetControlVisible ('T_NATUREAUXI_', False);
      SetControlVisible ('TT_NATUREAUXI', False);
      if pos ('CLI', Arguments) > 0 then
      begin
       SetControlText ('T_NATUREAUXI', 'CLI');
       Ecran.Caption:='Activation des clients' ;
      end else
      begin
       SetControlText ('T_NATUREAUXI', 'FOU');
       Ecran.Caption:='Activation des Fournisseurs' ;
      end;
   end;
   Ecran.HelpContext:=110000105 ;
   end
else begin
     bFermer:=true;
     SetControlText('T_FERME','-');
     if pos('GRC',Arguments)>0 then
         Ecran.Caption:='Mise en sommeil des clients/prospects'
     else
     begin
      SetControlVisible ('T_NATUREAUXI_', False);
      SetControlVisible ('TT_NATUREAUXI', False);
      if pos ('CLI', Arguments) > 0 then
      begin
         SetControlText ('T_NATUREAUXI', 'CLI');
         Ecran.Caption:='Mise en sommeil des clients' ;
      end else
      begin
         SetControlText ('T_NATUREAUXI', 'FOU');
         Ecran.Caption:='Mise en sommeil des Fournisseurs' ;
      end;
     end;
     Ecran.HelpContext:=110000085 ;
     end;
UpdateCaption(ecran) ;
if pos ('FOU', Arguments) > 0 then
   begin
   SetControlVisible ('PTABLELIBRE',False);
   SetControlVisible ('PZONELIBRE',False);
   SetControlVisible ('T_ENSEIGNE',False);
   SetControlVisible ('TT_ENSEIGNE',False);
   SetControlVisible ('T_ZONECOM',False);
   SetControlVisible ('TT_ZONECOM',False);
   SetControlVisible ('T_SOCIETEGROUPE',False);
   SetControlVisible ('TT_SOCIETEGROUPE',False);
   SetControlVisible ('T_PRESCRIPTEUR',False);
   SetControlVisible ('TT_PRESCRIPTEUR',False);
   SetControlVisible ('T_REPRESENTANT',False);
   SetControlVisible ('TT_REPRESENTANT',False);
   end;
if CtxScot in V_PGI.PGIContexte then // mcd 24/05/04 11001
 begin
 SetcontrolVisible ('T_ZONECOM',False);
 SetcontrolVisible ('TT_ZONECOM',False);
 SetcontrolVisible ('T_PRESCRIPTEUR',False);
 SetcontrolVisible ('TT_PRESCRIPTEUR',False);
 SetcontrolVisible ('T_REPRESENTANT',False);
 SetcontrolVisible ('TT_REPRESENTANT',False);
 end;
end;

procedure tof_gcouvferm.OnLoad;
var xx_where,Confid : string;
begin
inherited ;
{$IFDEF GRC}
  if (ctxGRC in V_PGI.PGIContexte) then
  begin
  if ( GetControlText ('T_NATUREAUXI') <> 'FOU') then Confid:='MOD' else Confid:='MODF';
  if (GetControlText('XX_WHERE') = '') then
      SetControlText('XX_WHERE',RTXXWhereConfident(Confid))
  else
      begin
      xx_where := GetControlText('XX_WHERE');
      xx_where := xx_where + RTXXWhereConfident(Confid);
      SetControlText('XX_WHERE',xx_where) ;
      end;
  end;
{$ENDIF}
end;

procedure tof_gcouvferm.OuvreFerme;
var  F : TFMul;
{$IFDEF EAGLCLIENT}
       L : THGrid;
{$ELSE}
       L : THDBGrid;
{$ENDIF}
     i : integer;
     Pages:TPageControl;
begin
F:=TFMul(Ecran);
if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
   begin
   PGIInfo('Aucun élément sélectionné','');
   exit;
   end;
if PGIAsk('Confirmez vous le traitement ?','')<>mrYes then exit ;

L:= F.FListe;
Pages:=F.Pages;

if L.AllSelected then
   begin
   stWhere:=RecupWhereCritere(Pages);
   stWhere:=stWhere+' AND not exists (SELECT AFF_AFFAIRE FROM AFFAIRE WHERE AFF_TIERS=T_TIERS '
         +' AND AFF_DATEDEBGENER <> "'+ UsDateTime(iDate1900)  +'" AND AFF_GENERAUTO <> "MAN" AND AFF_DATEFINGENER > "' + UsDateTime(V_PGI.DateEntree)+  '") ';
   if (Transactions (SetAllOuvreFerme,3) <> oeOK) then PGIBox('Impossible d''ouvrir ou fermer tous les tiers ', F.Caption);
   L.AllSelected:=False;
   end else
   begin
   InitMove(L.NbSelected,'');
   for i:=0 to L.NbSelected-1 do
      begin
      MoveCur(False);
      L.GotoLeBookmark(i);
{$IFDEF EAGLCLIENT}
      F.Q.TQ.Seek(F.FListe.Row-1) ;
{$ENDIF}
      CodeTiers:=F.Q.FindField('T_TIERS').asstring ;
      if bFermer and TiersAvecAffaireActive(CodeTiers) then
          begin
          if CtxScot in V_PGI.PGIContexte then PGIBox('Impossible de fermer le tiers '+CodeTiers+' : missions en cours.', F.Caption)
          else PGIBox('Impossible de fermer le tiers '+CodeTiers+' : contrat en cours.', F.Caption) ;
          end
        else if (Transactions (SetOuvreFerme,3) <> oeOK) then PGIBox('Impossible d''ouvrir ou fermer le tiers '+CodeTiers, F.Caption);
      end;
   L.ClearSelected;
   end;
FiniMove;
end ;


procedure tof_gcouvferm.SetAllOuvreFerme  ;
begin
{
if bfermer then
     ExecuteSql('UPDATE TIERS SET T_FERME="X", T_DATEFERMETURE="'+UsDateTime(Date)+'", '+
      'T_DATEMODIF="'+UsDateTime(Date)+'" '+StWhere)
else ExecuteSql('UPDATE TIERS SET T_FERME="-", T_DATEOUVERTURE="'+UsDateTime(Date)+'", '+
      'T_DATEMODIF="'+UsDateTime(Date)+'" '+StWhere);
}
if bfermer then
     ExecuteSql('UPDATE TIERS SET T_FERME="X", T_UTILISATEUR="'+V_PGI.User+'", T_DATEFERMETURE="'+UsDateTime(Date)+'", '+
      'T_DATEMODIF="'+UsTime(NowH)+'" where t_tiers in ( select t_tiers from RTTIERS '+StWhere+')')
else ExecuteSql('UPDATE TIERS SET T_FERME="-", T_UTILISATEUR="'+V_PGI.User+'", T_DATEOUVERTURE="'+UsDateTime(Date)+'", '+
      'T_DATEMODIF="'+UsTime(NowH)+'" where t_tiers in ( select t_tiers from RTTIERS '+StWhere+')');
end;

procedure tof_gcouvferm.SetOuvreFerme  ;
begin
if bfermer then
     ExecuteSql('UPDATE TIERS SET T_FERME="X", T_UTILISATEUR="'+V_PGI.User+'", T_DATEFERMETURE="'+UsDateTime(Date)+'", '+
       'T_DATEMODIF="'+UsTime(NowH)+'" Where T_Tiers="'+CodeTiers+'"')
else ExecuteSql('UPDATE TIERS SET T_FERME="-", T_UTILISATEUR="'+V_PGI.User+'", T_DATEOUVERTURE="'+UsDateTime(Date)+'", '+
       'T_DATEMODIF="'+UsTime(NowH)+'" Where T_Tiers="'+CodeTiers+'"');
end;


procedure AGLGCOuvreFermeTiers(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then TOTOF:=TFMul(F).LaTOF else exit;
if (TOTOF is tof_gcouvferm) then tof_gcouvferm(TOTOF).OuvreFerme else exit;
end;

Initialization
registerclasses([tof_gcouvferm]);
RegisterAglProc('GCOuvreFermeTiers',True,1,AGLGCOuvreFermeTiers);

end.
