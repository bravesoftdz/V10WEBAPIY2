{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 29/03/2016
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTDATEIMPUTATION ()
Mots clefs ... : TOF;BTDATEIMPUTATION
*****************************************************************}
Unit BTDATEIMPUTATION ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes,
     Vierge,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     //mul,
{$else}
     //eMul,
     uTob,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HTB97,
     UTOF ;

Type
  TOF_BTDATEIMPUTATION = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    BValide : TToolbarButton97;
    BAnnule : TToolbarButton97;
    //
    DateImputation : THEdit;
    //
    procedure AnnulerOnClick(Sender: TObject);
    procedure GetObjects;
    procedure SetScreenEvents;
    procedure ValiderOnClick(Sender: TObject);
  end ;

Implementation

procedure TOF_BTDATEIMPUTATION.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTDATEIMPUTATION.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTDATEIMPUTATION.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTDATEIMPUTATION.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTDATEIMPUTATION.OnArgument (S : String ) ;
begin
  Inherited ;
  //
  //Chargement des zones ecran dans des zones programme
  GetObjects;
  //
 //Gestion des évènement des zones écran
  SetScreenEvents;

  DateImputation.text := DateToStr(Now);

end ;

procedure TOF_BTDATEIMPUTATION.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTDATEIMPUTATION.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTDATEIMPUTATION.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTDATEIMPUTATION.GetObjects;
begin

  if Assigned(GetControl('BValider'))  then BValide   := TToolBarButton97(Getcontrol('BValider'));
  if Assigned(GetControl('BAnnuler'))  then BAnnule   := TToolBarButton97(Getcontrol('BAnnuler'));

  if Assigned(GetControl('DateImputation')) then DateImputation := THEdit(GetControl('DateImputation'));

end;

procedure TOF_BTDATEIMPUTATION.SetScreenEvents;
begin

  if Assigned(BValide)               then BValide.OnClick  := ValiderOnClick;
  if Assigned(BAnnule)               then Bannule.OnClick  := AnnulerOnClick;

end;

procedure TOF_BTDATEIMPUTATION.ValiderOnClick(Sender : TObject);
begin

   if LaTOB = nil then Exit;

   LaTOB.PutValue('DATEIMPUTATION', DateImputation.text);

   TFVierge(Ecran).Close;


end;
procedure TOF_BTDATEIMPUTATION.AnnulerOnClick(Sender : TObject);
begin

  LaTOB.ClearDetail;

end;

Initialization
  registerclasses ( [ TOF_BTDATEIMPUTATION ] ) ; 
end.

