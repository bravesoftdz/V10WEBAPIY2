{***********UNITE*************************************************
Auteur  ...... : EPI
Créé le ...... : 20/03/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PROCESSUSDETAIL ()
Mots clefs ... : TOF;PROCESSUSDETAIL
*****************************************************************}
{
PT1 14/04/2006 V_65 EPI FQ 12791 gestion du détail
PT2 13/09/2007 V_80 FCO FQ 14388 Concepts
}
Unit PROCESSUSDETAIL_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}  mul, FE_Main, DBGrids, Hdb,
{$else}
     eMul, MainEAGL,
{$ENDIF}
     forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox,
     HTB97,UTOF, UTOM,UTOB, Grids, Hqry;

Type
  TOF_PROCESSUSDETAIL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private
    predef,dossier,code,protype      : string;
    LectureSeule, CEG, STD, DOS : boolean;
{$IFDEF EAGLCLIENT}
  Grille: THGrid;
{$ELSE}
  Grille: THDBGrid;
{$ENDIF}
    procedure BTnInsClick(Sender: TObject);
    procedure DblClickGrill(Sender: TObject);
  end ;

Implementation
uses PgOutils2, Pgoutils;

procedure TOF_PROCESSUSDETAIL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PROCESSUSDETAIL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PROCESSUSDETAIL.OnUpdate ;
begin
  Inherited ;
 end ;

procedure TOF_PROCESSUSDETAIL.OnLoad ;
begin
  Inherited ;

  SetControlVisible('BInsert',False);
  LectureSeule := FALSE;
  AccesPredefini('TOUS', CEG, STD, DOS);
  if (predef = 'CEG') then
  begin
    LectureSeule := (CEG = False);
    //SetControlVisible('BInsert',LectureSeule);
  end;
  if (predef = 'STD') then
  begin
    LectureSeule := (STD = False);
    //SetControlVisible('BInsert',LectureSeule);
  end;
  if (predef = 'DOS') then
  begin
    LectureSeule := (DOS = False); //PT2
  end;
  If LectureSeule = FALSE then
    SetControlVisible('BInsert',True);

end;

procedure TOF_PROCESSUSDETAIL.OnArgument (S : String ) ;
var
  Arg,Ordre : String;
  Tit : String;                         // PT1
  BtnIns : TToolBarButton97;
begin
  Inherited ;
   Arg := S;
   predef:=Trim(ReadTokenPipe(Arg,';')) ;
   dossier:=Trim(ReadTokenPipe(Arg,';')) ;
   code := Trim(ReadTokenPipe(Arg,';')) ;
   protype := Trim(ReadTokenPipe(Arg,';')) ;
   // PT1   début ajout titre écran
   Tit := Trim(ReadTokenPipe(Arg,';')) ;
   TFMul(Ecran).Caption := 'Détail du processus '+ Code + ' - '+  Tit;
   // PT1 fin

   Ordre := ' PRD_PREDEFINI="';
   Ordre := Ordre + predef + '" AND PRD_NODOSSIER="';
   Ordre := Ordre + dossier + '" AND PRD_CODEPROC="';
   Ordre := Ordre + code + '"';
   SetControlText('XX_WHERE',Ordre);

  SetControlVisible('BINSERT',True);
  BtnIns := TToolBarButton97(GetControl('BINSERT'));
  if BtnIns <> nil then BtnIns.OnClick := BTnInsClick;

{$IFDEF EAGLCLIENT}
  Grille := THGrid(GetControl('Fliste'));
{$ELSE}
  Grille := THDBGrid(GetControl('Fliste'));
{$ENDIF}
   if Grille <> nil then
     Grille.OnDblClick := DblClickGrill;
end ;

procedure TOF_PROCESSUSDETAIL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PROCESSUSDETAIL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PROCESSUSDETAIL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_PROCESSUSDETAIL.BTnInsClick(Sender: TObject);
var
 St : String;
begin
  St :=   Predef + ';' + Dossier + ';' + Code + ';'+ ' ;'+ Protype + ';';
  // Sel := St;
  AGLLanceFiche('PAY','PROCESSUSDET','','','ACTION=CREATION;'+St);
  TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

procedure TOF_PROCESSUSDETAIL.DblClickGrill(Sender: TObject);
var
  ordre : integer;
  Sel,st : string;
begin
  sel := '';
if (Grille <> nil) and (TFMUL(Ecran).Q.RecordCount = 0)  then
    exit;

if (Grille <> nil) then
   begin
      predef := TFMUL(Ecran).Q.Findfield('PRD_PREDEFINI').asstring;
      dossier := TFMUL(Ecran).Q.FindField ('PRD_NODOSSIER').Asstring;
      code := TFMUL(Ecran).Q.FindField ('PRD_CODEPROC').Asstring;
      ordre := TFMUL(Ecran).Q.FindField ('PRD_NOORDRE').AsInteger;
      st := predef + ';'+ dossier + ';' + code + ';' +  Inttostr(ordre);
      Sel := St;
      St := St + ';' + Protype;
   end;
  AGLLanceFiche ('PAY','PROCESSUSDET',  '',Sel,'ACTION=MODIFICATION;'+ St);
  TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

Initialization
  registerclasses ( [ TOF_PROCESSUSDETAIL ] ) ;
end.
