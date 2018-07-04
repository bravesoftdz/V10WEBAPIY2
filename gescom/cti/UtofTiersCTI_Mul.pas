unit UTofTiersCti_Mul;

interface
uses UTOF,
{$IFDEF EAGLCLIENT}
     eMul,MaineAGL,
{$ELSE}
     Mul, Fe_Main,
{$ENDIF}
     Forms,Classes,M3FP,HMsgBox,HEnt1,HCtrls,UtilSelection,UtilRT
     ,SysUtils,ParamSoc, stdctrls ;

Function RTLanceFiche_TiersCti_Mul(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Type
     TOF_TiersCti_Mul = Class(TOF)
       private
       procedure AppelleLesTiers;
       public
       Argument,Origine,WhereUpdate : string;
       procedure OnLoad ; override ;
       procedure OnClose ; override ;
       procedure OnArgument (stArgument : string); override;
     end;

procedure AGLRTCTI_AppelleLesTiers(parms : array of variant; nb : integer);

implementation

Function RTLanceFiche_TiersCti_Mul(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_TiersCti_Mul.OnArgument (stArgument : string);
var F : TForm;
begin
F := TForm (Ecran);
MulCreerPagesCL(F,'NOMFIC=GCTIERS');
//MulCreerPagesCL(F);
end;

procedure TOF_TiersCti_Mul.OnLoad;
var xx_where,StWhere,StListeActions: string;
    DateDeb,DateFin : TDateTime;
begin
StWhere := '';
WhereUpdate:='';
if (GetControlText('SANSSELECT') <> 'X') then
   begin
   DateDeb := StrToDate(GetControlText('DATEACTION'));
   DateFin := StrToDate(GetControlText('DATEACTION_'));
   if GetControlText('SANS') = 'X' then StWhere := 'NOT ' ;
   StWhere:=StWhere + 'EXISTS (SELECT RAC_NUMACTION FROM ACTIONS WHERE (RAC_AUXILIAIRE = RTTIERS.T_AUXILIAIRE';
   if GetControlText('TYPEACTION') <> '' then
      begin
      StListeActions:=FindEtReplace(GetControlText('TYPEACTION'),';','","',True);
      StListeActions:='("'+copy(StListeActions,1,Length(StListeActions)-2)+')';
      StWhere:=StWhere + ' AND RAC_TYPEACTION in '+StListeActions;
      // pour Update Action existante
      if GetControlText('AVEC') = 'X' then
          WhereUpdate:=' AND RAC_TYPEACTION in '+StListeActions;

      end;
   if GetControlText('ETATACTION') <> '' then
      begin
      StListeActions:=FindEtReplace(GetControlText('ETATACTION'),';','","',True);
      StListeActions:='("'+copy(StListeActions,1,Length(StListeActions)-2)+')';
      StWhere:=StWhere + ' AND RAC_ETATACTION in '+StListeActions;
      // pour Upadte Action existante
      if GetControlText('AVEC') = 'X' then
          WhereUpdate:=whereUpdate+' AND RAC_ETATACTION in '+StListeActions;
      end;



   StWhere := StWhere + ' AND RAC_DATEACTION >= "'+UsDateTime(DateDeb) +'" AND RAC_DATEACTION <= "'+UsDateTime(DateFin)+'"))';
   // pour Upadte Action existante
   if GetControlText('AVEC') = 'X' then
        WhereUpdate:=whereUpdate+' AND RAC_DATEACTION >= "'+UsDateTime(DateDeb) +'" AND RAC_DATEACTION <= "'+UsDateTime(DateFin)+'"';

   // on ne s'encombre pas si le ParamSoc n'est pas positif
   if (not GetParamSoc('SO_RTCTIMODACTEFFECT')) then WhereUpdate:='';

   if (GetControlText('XX_WHERE') = '') then
      SetControlText('XX_WHERE',StWhere)
   else
       begin
       xx_where := GetControlText('XX_WHERE');
       xx_where := xx_where + ' and (' + StWhere + ')';
       SetControlText('XX_WHERE',xx_where) ;
       end;
   end
else
   // dans le cas de modification de l'action d'origine, la sélection de l'onglet
   // Action est prioritaire par rapport à celui de l'action d'origine
   if (GetParamSoc('SO_RTCTIMODACTEFFECT')) and (GetCheckBoxState('EXISTE')=cbChecked) then
      WhereUpdate:='AND RAC_OPERATION="'+GetControlText('OPERATION')+'" and RAC_NUMACTGEN='+GetControlText('NUMACTGEN');

if (GetControlText('XX_WHERE') = '') then
    SetControlText('XX_WHERE',RTXXWhereConfident('CON'))
else
    begin
    xx_where := GetControlText('XX_WHERE');
    xx_where := xx_where + RTXXWhereConfident('CON');
    SetControlText('XX_WHERE',xx_where) ;
    end;



end;

procedure TOF_TiersCti_Mul.OnClose;
begin
end;

/////////////// AppelleLesTiers //////////////
procedure TOF_TiersCti_Mul.AppelleLesTiers;
Var F : TFMul ;
    StArg,StArg2,Cle,Retour,stTiers : String;
    i : integer;
begin
F:=TFMul(Ecran);
if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
  begin MessageAlerte('Aucun élément sélectionné'); exit; end;
{$IFDEF EAGLCLIENT}
if F.bSelectAll.Down then
   if not F.Fetchlestous then
     begin
     F.bSelectAllClick(Nil);
     F.bSelectAll.Down := False;
     exit;
     end else
     begin
     F.bSelectAll.Down := True;
     F.Fliste.AllSelected := true;
     end;
{$ENDIF}

if F.FListe.AllSelected then
    begin
    F.Q.DisableControls ;
    F.Q.First ;
    While not F.Q.Eof do
        begin
        Cle := F.Q.FindField('T_AUXILIAIRE').AsString;
        stArg:=''; stArg2:='';

        stTiers:= F.Q.FindField('T_TIERS').AsString;
        if (RTDroitModiftiers(stTiers)=False) then
            stArg:= 'ACTION=CONSULTATION;';
        if WhereUpdate<> '' then
           stArg2:=';UPDCTI='+ WhereUpdate;

        Retour:=AGLlancefiche('GC','GCTIERS','',Cle,'NUMEROTER;SERIECTI;MONOFICHE;'+stArg+'T_NATUREAUXI='+F.Q.FindField('T_NATUREAUXI').AsString+stArg2);
        if Retour='Stop' then Break;
        //GenereAction(Retour,stTiers,Cle);
        F.Q.Next ;
        end;
    end
  else
    begin
    for i:=0 to F.FListe.NbSelected-1 do
        begin
        F.FListe.GotoLeBookmark(i) ;
{$IFDEF EAGLCLIENT}
        F.Q.TQ.Seek(F.FListe.Row-1) ;
{$ENDIF}
        Cle := F.Q.FindField('T_AUXILIAIRE').AsString;
        stArg:=''; stArg2:='';
        stTiers:= F.Q.FindField('T_TIERS').AsString;
        if (RTDroitModiftiers(stTiers)=False) then
            stArg:= 'ACTION=CONSULTATION;';
        if WhereUpdate<> '' then
           stArg2:=';UPDCTI='+ WhereUpdate;
        Retour:=AGLlancefiche('GC','GCTIERS','',Cle,'NUMEROTER;SERIECTI;MONOFICHE;'+stArg+'T_NATUREAUXI='+F.Q.FindField('T_NATUREAUXI').AsString+stArg2);
        if Retour='Stop' then Break;
        //GenereAction(Retour,stTiers,Cle);
        end;
    end;
    F.Q.EnableControls;

if F.bSelectAll.Down then
    begin
    F.bSelectAllClick(Nil);
    F.bSelectAll.Down := False;
    end;

end;


/////////////// Procedure appellé par le bouton Validation //////////////
procedure AGLRTCTI_AppelleLesTiers(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then TOTOF:=TFMul(F).LaTOF else exit;
if (TOTOF is TOF_TiersCti_Mul) then TOF_TiersCti_Mul(TOTOF).AppelleLesTiers else exit;
end;

Initialization
registerclasses([TOF_TiersCti_Mul]) ;
RegisterAglProc('RTCTI_AppelleLesTiers',TRUE,0,AGLRTCTI_AppelleLesTiers);
end.
