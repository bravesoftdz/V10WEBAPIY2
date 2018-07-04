{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 09/04/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFRES_SYNCWHAT ()
Mots clefs ... : TOF;AFRES_SYNCWHAT
*****************************************************************}
Unit UTofAFRes_SyncWhat ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ELSE}
     utob,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     Dicobtp, 
     UTOF ; 

Type
  TOF_AFRES_SYNCWHAT = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

    procedure Coche_OnClick(Sender : TObject);

  private
    DontClose : boolean;
  end ;

Implementation
uses Vierge;

procedure TOF_AFRES_SYNCWHAT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFRES_SYNCWHAT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFRES_SYNCWHAT.OnUpdate ;
var S : String;
begin
  if (GetControlText('CHK_COMMONFIELDS') <> 'X') and (GetControlText('CHK_COST') <> 'X') and (GetControlText('CHK_RATE') <> 'X') then
  begin
//     LastError := 1;
     PGIBoxAF('Merci de choisir au moins un critère d''alignement', '');
     DontClose := true;
     exit;
  end;

  S := 'COMMONFIELDS='+GetControlText('CHK_COMMONFIELDS')+';'+
       'ALIMCOST='+GetControlText('CHK_COST')+';';

  if GetControlText('CHK_COST') = 'X' then
  begin
     S := S+'COSTFROM=';
     if GetControlText('RB_COST_THO') = 'X' then S := S+'THO;';
     if GetControlText('RB_COST_THI') = 'X' then S := S+'THI,'+GetControlText('ED_COST_THD')+','+GetControlText('ED_COST_THD_')+','+GetControlText('CHK_COST_THC')+';';
     if GetControlText('RB_COST_TMO') = 'X' then S := S+'TMO,'+GetControlText('ED_COST_TMD')+','+GetControlText('ED_COST_TMD_')+','+GetControlText('CHK_COST_TMC')+';';
  end;

  S := S+'ALIMRATE='+GetControlText('CHK_RATE')+';';
  if GetControlText('CHK_RATE') = 'X' then
  begin
     S := S+'RATEFROM=';
     if GetControlText('RB_RATE_THO') = 'X' then S := S+'THO;';
     if GetControlText('RB_RATE_THI') = 'X' then S := S+'THI,'+GetControlText('ED_RATE_THD')+','+GetControlText('ED_RATE_THD_')+','+GetControlText('CHK_RATE_THC')+';';
     if GetControlText('RB_RATE_TMO') = 'X' then S := S+'TMO,'+GetControlText('ED_RATE_TMD')+','+GetControlText('ED_RATE_TMD_')+','+GetControlText('CHK_RATE_TMC')+';';
  end;

  S := S+'HISTO='+GetControlText('CHK_HISTO')+','+GetControlText('ED_HISTODATE')+';';

  TFVierge(Ecran).Retour := S;
  Inherited ;
end ;

procedure TOF_AFRES_SYNCWHAT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_AFRES_SYNCWHAT.OnArgument (S : String ) ;
var Q : TQuery;
begin
  Inherited ;
  DontClose := false;

  SetControlText('ED_COST_THD', FormatDateTime('dd"/"mm"/"yyyy', DebutDeMois(PlusMois(Now, -1))));
  SetControlText('ED_COST_THD_', FormatDateTime('dd"/"mm"/"yyyy', FinDeMois(PlusMois(Now, -1))));
  SetControlText('ED_RATE_THD', FormatDateTime('dd"/"mm"/"yyyy', DebutDeMois(PlusMois(Now, -1))));
  SetControlText('ED_RATE_THD_', FormatDateTime('dd"/"mm"/"yyyy', FinDeMois(PlusMois(Now, -1))));

  Q := OpenSQL('SELECT PEX_DATEDEBUT, PEX_DATEFIN FROM EXERSOCIAL ORDER BY PEX_DATEFIN DESC', true);
  if Q.EOF then
  begin
     SetControlText('ED_COST_TMD', FormatDateTime('"01/01/"yyyy', Now-365));
     SetControlText('ED_COST_TMD_', FormatDateTime('"31/12/"yyyy', Now-365));
     SetControlText('ED_RATE_TMD', FormatDateTime('"01/01/"yyyy', Now-365));
     SetControlText('ED_RATE_TMD_', FormatDateTime('"31/12/"yyyy', Now-365));
  end else
  begin
     SetControlText('ED_COST_TMD', FormatDateTime('dd"/"mm"/"yyyy', Q.FindField('PEX_DATEDEBUT').AsDateTime));
     SetControlText('ED_COST_TMD_', FormatDateTime('dd"/"mm"/"yyyy', Q.FindField('PEX_DATEFIN').AsDateTime));
     SetControlText('ED_RATE_TMD', FormatDateTime('dd"/"mm"/"yyyy', Q.FindField('PEX_DATEDEBUT').AsDateTime));
     SetControlText('ED_RATE_TMD_', FormatDateTime('dd"/"mm"/"yyyy', Q.FindField('PEX_DATEFIN').AsDateTime));
  end;
  Ferme(Q);

  TCheckBox(GetControl('CHK_COST')).OnClick := Coche_OnClick;
  TRadioButton(GetControl('RB_COST_THO')).OnClick := Coche_OnClick;
  TRadioButton(GetControl('RB_COST_THI')).OnClick := Coche_OnClick;
  TRadioButton(GetControl('RB_COST_TMO')).OnClick := Coche_OnClick;
  TCheckBox(GetControl('CHK_RATE')).OnClick := Coche_OnClick;
  TRadioButton(GetControl('RB_RATE_THO')).OnClick := Coche_OnClick;
  TRadioButton(GetControl('RB_RATE_THI')).OnClick := Coche_OnClick;
  TRadioButton(GetControl('RB_RATE_TMO')).OnClick := Coche_OnClick;
  TCheckBox(GetControl('CHK_HISTO')).OnClick := Coche_OnClick;
  Coche_OnClick(nil);
end ;

procedure TOF_AFRES_SYNCWHAT.OnClose ;
begin
  Inherited ;
  if DontClose then
  begin
     LastError := -1;
     LastErrorMsg := '';
     DontClose := false;
     exit;     // il y eu une erreur, on reste sur la même fiche
  end;
end ;

procedure TOF_AFRES_SYNCWHAT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AFRES_SYNCWHAT.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_AFRES_SYNCWHAT.Coche_OnClick(Sender : TObject);
begin
   SetControlEnabled('ED_COST_THD', (GetControlText('RB_COST_THI') = 'X') and (GetControlText('CHK_COST') = 'X'));
   SetControlEnabled('ED_COST_THD_', (GetControlText('RB_COST_THI') = 'X') and (GetControlText('CHK_COST') = 'X'));
   SetControlEnabled('CHK_COST_THC', (GetControlText('RB_COST_THI') = 'X') and (GetControlText('CHK_COST') = 'X'));

   SetControlEnabled('ED_COST_TMD', (GetControlText('RB_COST_TMO') = 'X') and (GetControlText('CHK_COST') = 'X'));
   SetControlEnabled('ED_COST_TMD_', (GetControlText('RB_COST_TMO') = 'X') and (GetControlText('CHK_COST') = 'X'));
   SetControlEnabled('CHK_COST_TMC', (GetControlText('RB_COST_TMO') = 'X') and (GetControlText('CHK_COST') = 'X'));

   SetControlEnabled('ED_RATE_THD', (GetControlText('RB_RATE_THI') = 'X') and (GetControlText('CHK_RATE') = 'X'));
   SetControlEnabled('ED_RATE_THD_', (GetControlText('RB_RATE_THI') = 'X') and (GetControlText('CHK_RATE') = 'X'));
   SetControlEnabled('CHK_RATE_THC', (GetControlText('RB_RATE_THI') = 'X') and (GetControlText('CHK_RATE') = 'X'));

   SetControlEnabled('ED_RATE_TMD', (GetControlText('RB_RATE_TMO') = 'X') and (GetControlText('CHK_RATE') = 'X'));
   SetControlEnabled('ED_RATE_TMD_', (GetControlText('RB_RATE_TMO') = 'X') and (GetControlText('CHK_RATE') = 'X'));
   SetControlEnabled('CHK_RATE_TMC', (GetControlText('RB_RATE_TMO') = 'X') and (GetControlText('CHK_RATE') = 'X'));

   SetControlEnabled('ED_HISTODATE', (GetControlText('CHK_HISTO') = 'X'));
end;

Initialization
  registerclasses ( [ TOF_AFRES_SYNCWHAT ] ) ;
end.

