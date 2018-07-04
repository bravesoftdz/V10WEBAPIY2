{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 20/11/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : ESPCPQRIRPF115 ()
Mots clefs ... : TOF;ESPCPQRIRPF115
*****************************************************************}
Unit ESPCPQRIRPF115_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
     uTob,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     TofMeth ;

procedure CP_LanceFicheESPCPQRMOD347;
procedure CP_LanceFicheESPCPQRJALTVA(szType : String);
procedure CP_LanceFicheESPCPQRTVA300;
procedure CP_LanceFicheESPCPQRIRPF110;
procedure CP_LanceFicheESPCPQRIRPF115;

Type
  TOF_ESPCPQROfficial = Class (TOF_Meth)
   private
    EXERCICE        : THValComboBox ;
    DATECOMPTABLE,
    DATECOMPTABLE_  : THedit ;
    DateD,DateF       : TDatetime ;
    procedure ExoOnChange(Sender : TObject);
    procedure DateOnExit(Sender : TObject) ;
   public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    {Constructor} Procedure  TRUC_create ; //override ;
  end ;

  TOF_ESPCPQRTVA300 = Class (TOF_ESPCPQROfficial)
   public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

   TOF_ESPCPQRJALTVA = Class (TOF_METH)
   public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

  TOF_ESPCPQRIRPF110 = Class (TOF_ESPCPQROfficial)
   public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

  TOF_ESPCPQRIRPF115 = Class (TOF_ESPCPQROfficial)
   public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

Implementation
uses ULibExercice, Ent1, Menus,
{$IFDEF eAGLClient}
 eQRS1, MaineAGL
{$ELSE}
 QRS1, FE_MAIN
{$ENDIF}
 ;

procedure CP_LanceFicheESPCPQRMOD347;
begin
  AGLLanceFiche('CP','ESPCPQRMOD347','','','');
end;

procedure CP_LanceFicheESPCPQRJALTVA(szType : String);
begin
  AGLLanceFiche('CP','ESPCPQRJALTVA','','',szType);
end;

procedure CP_LanceFicheESPCPQRTVA300;
begin
  AGLLanceFiche('CP','ESPCPQRTVA300','','','');
end;

procedure CP_LanceFicheESPCPQRIRPF110;
begin
  AGLLanceFiche('CP','ESPCPQRIRPF110','','','');
end;

procedure CP_LanceFicheESPCPQRIRPF115;
begin
  AGLLanceFiche('CP','ESPCPQRIRPF115','','','');
end;

//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
Procedure TOF_ESPCPQROfficial.TRUC_create ;
Begin
  // inherited ;
  DATECOMPTABLE:=THEdit(GetControl('DATECOMPTABLE')) ; DATECOMPTABLE.OnExit:=DateOnExit ;

  DATECOMPTABLE_:=THEdit(GetControl('DATECOMPTABLE_')) ; DATECOMPTABLE_.OnExit:= DateOnExit ;

  EXERCICE:= THValComboBox(GetControl('EXERCICE')) ;
  CInitComboExercice(EXERCICE);
  EXERCICE.OnChange:=ExoOnChange ;
  if (CtxPCl in V_PGI.PgiContexte) and  (VH^.CPExoRef.Code<>'') then
     EXERCICE.Value:=CExerciceVersRelatif(VH^.CPExoRef.Code)
  else
     EXERCICE.Value:=CExerciceVersRelatif(VH^.Entree.Code) ;

End ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQROfficial.OnNew ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQROfficial.OnDelete ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQROfficial.OnUpdate ;
begin
  Inherited ;
  if trim(GetControlText('FFiltres'))='' then
     TMenuItem(GetControl('BCreerFiltre')).Click ;
  TMenuItem(GetControl('BSaveFiltre')).Click ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQROfficial.OnLoad ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQROfficial.OnArgument (S : String ) ;
begin
  Truc_Create ;
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQROfficial.OnClose ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQROfficial.OnDisplay () ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQROfficial.OnCancel () ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQROfficial.ExoOnChange(Sender: TObject);
Begin
  if (Assigned(EXERCICE) and (Assigned(DATECOMPTABLE)) and (Assigned(DATECOMPTABLE_))) then
    CExoRelatifToDates(EXERCICE.Value,DATECOMPTABLE,DATECOMPTABLE_);
  DateD:=StrToDate(DATECOMPTABLE.Text) ;
  DateF:=StrToDate(DATECOMPTABLE_.Text) ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQROfficial.DateOnExit(Sender: TObject);
begin
 DoDateOnExit(THEdit(Sender), DATECOMPTABLE, DATECOMPTABLE_, DateD, DateF);
end;
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRTVA300.OnNew ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRTVA300.OnDelete ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRTVA300.OnUpdate ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRTVA300.OnLoad ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRTVA300.OnArgument (S : String ) ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRTVA300.OnClose ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRTVA300.OnDisplay () ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRTVA300.OnCancel () ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRIRPF110.OnNew ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRIRPF110.OnDelete ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRIRPF110.OnUpdate ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRIRPF110.OnLoad ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRIRPF110.OnArgument (S : String ) ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRIRPF110.OnClose ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRIRPF110.OnDisplay () ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRIRPF110.OnCancel () ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRIRPF115.OnNew ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRIRPF115.OnDelete ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRIRPF115.OnUpdate ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRIRPF115.OnLoad ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRIRPF115.OnArgument (S : String ) ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRIRPF115.OnClose ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRIRPF115.OnDisplay () ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
procedure TOF_ESPCPQRIRPF115.OnCancel () ;
begin
  Inherited ;
end ;
//////////////////////////////////////////////////////////////////////////////////
{ TOF_ESPCPQRJALTVA }

procedure TOF_ESPCPQRJALTVA.OnArgument(S: String);
begin
  if (S = 'TYPEETAT=VENTE') then TFQRS1(ECRAN).CodeEtat := 'JTV'
  else TFQRS1(ECRAN).CodeEtat := 'JTA';
  inherited;
end;

procedure TOF_ESPCPQRJALTVA.OnCancel;
begin
  inherited;

end;

procedure TOF_ESPCPQRJALTVA.OnClose;
begin
  inherited;

end;

procedure TOF_ESPCPQRJALTVA.OnDelete;
begin
  inherited;

end;

procedure TOF_ESPCPQRJALTVA.OnDisplay;
begin
  inherited;

end;

procedure TOF_ESPCPQRJALTVA.OnLoad;
begin
  inherited;

end;

procedure TOF_ESPCPQRJALTVA.OnNew;
begin
  inherited;

end;

procedure TOF_ESPCPQRJALTVA.OnUpdate;
begin
  inherited;

end;

Initialization
  registerclasses ( [ TOF_ESPCPQRJALTVA ,TOF_ESPCPQROfficial,TOF_ESPCPQRTVA300,TOF_ESPCPQRIRPF110,TOF_ESPCPQRIRPF115 ] ) ;
end.
