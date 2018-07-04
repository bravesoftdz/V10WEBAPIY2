{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 03/01/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : ASSSECSTRUC ()
Mots clefs ... : TOF;ASSSECSTRUC
*****************************************************************}
Unit ASSSECSTRUC_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     dbtables,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     UTOB,
     Vierge,
     Grids,
     Ent1  ;

Type
  TOF_ASSSECSTRUC = Class (TOF)
    TobStruc : Tob ;
    GrilleSec : THGrid ;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  private
    Axe : string ;
    CodeSec, LibSec : String ;
    Indice : integer ;
    LongMax : integer ;
    procedure ChargeTobStruc;
    procedure GrilleSecDblClick(sender: TObject);
  end ;

Implementation

procedure TOF_ASSSECSTRUC.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_ASSSECSTRUC.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_ASSSECSTRUC.OnUpdate ;
begin
  Inherited ;
  // renvoi du code et du libellé
  TFVierge(Ecran).retour:= CodeSec + '|' + LibSec ;
end ;

procedure TOF_ASSSECSTRUC.OnLoad ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 09/01/2002
Modifié le ... :   /  /    
Description .. : Remplissage de la TOB avec les structures et sous 
Suite ........ : sections.
Mots clefs ... : SECTION;STRUCTURE
*****************************************************************}
procedure TOF_ASSSECSTRUC.ChargeTobStruc ;
var ReqSql : string ;
    Q : TQuery ;
    TobS : Tob ;
    idxStruc : integer ;
begin
  ReqSql := 'SELECT SS_AXE, SS_SOUSSECTION FROM STRUCRSE' ;
  ReqSql := ReqSql + ' WHERE SS_AXE="'+Axe+'"' ;
  ReqSql := ReqSql + ' ORDER BY SS_SOUSSECTION' ;
  Q := OpenSQL (ReqSql,True ) ;
  TobStruc.LoadDetailDB('STRUCRSE', '', '', Q, False, True ) ;
  Ferme(Q) ;

  If TobStruc.Detail.Count > 0 Then
  begin
  ReqSql := 'SELECT PS_AXE, PS_SOUSSECTION, PS_CODE, PS_LIBELLE FROM SSSTRUCR' ;
  ReqSql := ReqSql + ' WHERE PS_AXE="'+Axe+'"';
  ReqSql := ReqSql + ' ORDER BY PS_SOUSSECTION' ;
  Q:=OpenSql(ReqSql,True) ;
  If not Q.Eof then
    begin
    TobS:=TOB.Create('_STRUCT', Nil, -1) ;
    TobS.LoadDetailDB('SSSTRUCR', '', '', Q, True, True) ;
    For IdxStruc := 0 to TobStruc.Detail.Count -1 do
      begin
      While (TobS.Detail.Count > 0) and (TobStruc.Detail[IdxStruc].GetValue('SS_SOUSSECTION')=TobS.Detail[0].GetValue('PS_SOUSSECTION')) Do
         begin
           TobS.Detail[0].ChangeParent(TobStruc.Detail[IdxStruc], -1) ;
         end ;
      end ;
    TobS.free ;
    end ;
  Ferme(Q) ;
  end ;

end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 09/01/2002
Modifié le ... :   /  /    
Description .. : Au double-clic sur la grille, s'il y a un plan suivant, on 
Suite ........ : sauvegarde la sous section sélectionnée et on affiche le 
Suite ........ : plan suivant.
Mots clefs ... : 
*****************************************************************}
procedure TOF_ASSSECSTRUC.GrilleSecDblClick ( sender : TObject ) ;
begin
  if Length(CodeSec) < LongMax then
    begin
    CodeSec := CodeSec + GrilleSec.Cells[0,GrilleSec.Row] ;
    LibSec := LibSec + GrilleSec.Cells[1,GrilleSec.Row]+' ' ;
    SetControlText('CODSEC',CodeSec ) ;
    SetControlText('LIBSEC',LibSec ) ;
    Inc(Indice) ;
    end ;
  if Indice <= TobStruc.Detail.Count-1 then
    begin
    GrilleSec.VidePile(false) ;
    TobStruc.Detail[Indice].PutGridDetail(GrilleSec,false,false,'PS_CODE;PS_LIBELLE') ;
    end ;
end ;


procedure TOF_ASSSECSTRUC.OnArgument (S : String ) ;
begin
  Inherited ;
  GrilleSec := THGrid(GetControl('GRILLESEC')) ;

  TobStruc := Tob.Create('Structure',Nil,-1) ;
  Axe := S ;
  ChargeTobStruc ;
  Indice := 0;
  if TobStruc.Detail.Count > 0 then
    begin
    GrilleSec.OnDblClick := GrilleSecDblClick ;
    TobStruc.Detail[Indice].PutGridDetail(GrilleSec,false,false,'PS_CODE;PS_LIBELLE') ;
    end ;

  LongMax := VH^.Cpta[AxeToFb(Axe)].Lg ;
  SetControlProperty('CODSEC','MAXLENGTH',LongMax ) ;
  CodeSec := ''; LibSec := '' ;
  SetFocusControl('BVALIDER') ;
end ;

procedure TOF_ASSSECSTRUC.OnClose ;
begin
  TobStruc.Free ;
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_ASSSECSTRUC ] ) ;
end.
