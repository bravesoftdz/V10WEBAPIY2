{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 05/06/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : AFCHOIXDATESREVAL ()
Mots clefs ... : TOF;AFCHOIXDATESREVAL
*****************************************************************}
Unit UTOFAFCHOIXDATESREVAL ;

Interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
{$IFDEF EAGLCLIENT}
    Maineagl,
{$ELSE}
   {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} db, FE_Main,
{$ENDIF}
      HCtrls,UTOF,AffaireUtil,HEnt1,EntGC,HTB97,UTOB,Stat,M3FP,ActiviteUtil;
      

Type
  TOF_AFCHOIXDATESREVAL = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    private
    chkToutesDates      :TCheckBox;
    chkPR               :TCheckBox;
    chkPV               :TCheckBox;
    EditDateDebut       : THEdit;
    EditDateFin         : THEdit;
    EditDateDebutMax    : THEdit;
    EditDateFinMax      : THEdit;
    gDateDebutAct, gDateFinAct:TDateTime;
    procedure CHK_TOUTESDATESClick(Sender: TObject);

  end ;

Function AFLanceFiche_ChoixDatesRevalActiv( Arg:string ):variant;

Implementation

procedure TOF_AFCHOIXDATESREVAL.OnArgument (S : String ) ;
begin
  Inherited ;

chkPR:=TCheckBox(GetControl('CHK_PR'));
chkPV:=TCheckBox(GetControl('CHK_PV'));
chkToutesDates:=TCheckBox(GetControl('CHK_TOUTESDATES'));
if (chkToutesDates<>nil) then
    begin
    chkToutesDates.OnClick := CHK_TOUTESDATESClick;
    end;

IntervalleDatesActivite(gDateDebutAct, gDateFinAct);
EditDateDebut:=THEdit(GetControl('DATEACT'));
EditDateFin:=THEdit(GetControl('DATEACT_'));
EditDateDebutMax:=THEdit(GetControl('DATEDEBUTMAX'));
EditDateFinMax:=THEdit(GetControl('DATEFINMAX'));

if ( EditDateDebut<>nil) then EditDateDebut.Text := datetostr(gDateDebutAct);
if ( EditDateFin<>nil) then EditDateFin.Text := datetostr(gDateFinAct);
if ( EditDateDebutMax<>nil) then EditDateDebutMax.Text := datetostr(gDateDebutAct);
if ( EditDateFinMax<>nil) then EditDateFinMax.Text := datetostr(gDateFinAct);

(*if (S='RES') then
    begin
    if (VH_GC.AFValoActPR<>'RES') then
        begin
        chkPR.Checked := false;
        chkPR.Enabled := false;
        end;
    if (VH_GC.AFValoActPV<>'RES') then
        begin
        chkPV.Checked := false;
        chkPV.Enabled := false;
        end;
    if (VH_GC.AFValoActPR<>'RES') and (VH_GC.AFValoActPV<>'RES') then
        begin
        TToolBarButton97(GetControl('BValider')).Enabled:=false;
        THLabel(GetControl('LBLREVALIMPOSSIBLE')).visible:=true;
        THLabel(GetControl('LBLREVALIMPOSSIBLE1')).visible:=true;
        end;
    end;
    *)

// Gestion de la préinitialisation des cases à cocher PR et PV suivant le paramétrage
chkPR.Checked := false;
chkPV.Checked := false;
if (S='RES') then
    begin
    if (VH_GC.AFValoActPR='RES') then
        begin
        chkPR.Checked := true;
        end;
    if (VH_GC.AFValoActPV='RES') then
        begin
        chkPV.Checked := true;
        end;
    end
else
if (S='ART') then
    begin
    if (VH_GC.AFValoActPR='ART') then
        begin
        chkPR.Checked := true;
        end;
    if (VH_GC.AFValoActPV='ART') then
        begin
        chkPV.Checked := true;
        end;
    end;

//  Préinitialisation de la revalorisation du planning
  SetControlVisible ('GBPLAN',VH_GC.GAPlanningSeria);
  SetControlChecked ('CHK_Planning',False);
  SetControlText ('DATEPLA',datetostr(gDateDebutAct));
  SetControlText ('DATEPLA_',datetostr(gDateFinAct));
  if (S='PLANNING') then
    SetControlChecked ('CHK_Planning',VH_GC.GAPlanningSeria);

end ;


procedure TOF_AFCHOIXDATESREVAL.CHK_TOUTESDATESClick(Sender: TObject);
begin
if (chkToutesDates.Checked = true) then
	begin
	SetControlEnabled('LblDateDebut', false);
	SetControlEnabled('LblDateFin', false);
	SetControlEnabled('DATEACT', false);
	SetControlEnabled('DATEACT_', false);
    EditDateDebut.Text := datetostr(gDateDebutAct);
    EditDateFin.Text := datetostr(gDateFinAct);
	//DATEDEBUT.Text:='';
	//DATEFIN.Text:='';
	end
else
	begin
	SetControlEnabled('LblDateDebut', true);
	SetControlEnabled('LblDateFin', true);
	SetControlEnabled('DATEACT', true);
	SetControlEnabled('DATEACT_', true);
	end;

end;

procedure TOF_AFCHOIXDATESREVAL.OnClose ;
begin
  Inherited ;
end ;

Function AFLanceFiche_ChoixDatesRevalActiv( Arg:string ):variant;
begin
result:=AGLLanceFiche ('AFF','AFCHOIXDATESREVAL','','',Arg);
end;

Initialization
  registerclasses ( [ TOF_AFCHOIXDATESREVAL ] ) ;
end.
