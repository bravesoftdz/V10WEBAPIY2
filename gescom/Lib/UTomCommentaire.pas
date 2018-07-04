{***********UNITE*************************************************
Auteur  ...... : Dominique BROSSET
Créé le ...... : 14/03/2001
Modifié le ... : 24/09/2001
Description .. : TOM de la TABLE : COMMENTAIRE
Suite ........ : (COMMENTAIRE)
Mots clefs ... : COMMENTAIRE
*****************************************************************}
Unit UTomCommentaire ;

Interface

Uses StdCtrls, Controls, Classes,forms, sysutils, ComCtrls,
     HCtrls, HEnt1, HMsgBox, UTOM, UTob, Windows, Math,
{$IFDEF EAGLCLIENT}
      UtileAGL,eFiche,eFichList,
{$ELSE}
      DBCtrls,db,dbTables,FichList,Fiche, HDB,
{$ENDIF}
     HTB97, HRichOLE;

Type
  TOM_Commentaire = Class (TOM)
    GLigne : THGrid;

    // gestion des données
    procedure CreerTobLigne(Arow :integer);
    procedure Enregistrer;
    function  GetTobLigne(Arow :integer):TOB;
    procedure InitialiseGrille;
    procedure InitialiseLigne (ARow : integer);
    procedure RempliLesLignes;
    procedure SupprimeLigne (ARow : Longint) ;

    // gestion des actions
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure GLigneRowEnter (Sender: TObject; Ou: Integer;
                              var Cancel: Boolean; Chg: Boolean);
    procedure GLigneRowExit (Sender: TObject; Ou: Integer;
                             var Cancel: Boolean; Chg: Boolean);
    procedure GLigneExit (Sender: TObject);

    // procédures héritées
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnArgument (S : String )   ; override ;
    procedure OnClose                    ; override ;
  private
    TobLigCom : TOB;
    stLibelle : string;
    bPasRowExit : boolean;
    end ;

Const
    iColLibelle : integer = 0;

Implementation

// Gestion des données *********************************************************

Procedure TOM_Commentaire.CreerTOBLigne (ARow : integer);
begin
if (ARow <> 0) and (ARow <> TobLigCom.Detail.Count) then exit;
TOB.Create ('COMMENTAIRE', TobLigCom, ARow) ;
InitialiseLigne (ARow);
end;

procedure TOM_Commentaire.Enregistrer;
Var iRow, iNbRow : Integer;
    TLBContenu : TStringList;
begin
TLBContenu := TStringList.Create;
if (GetControlText ('GCT_CODE') <> '') and (GetControlText ('GCT_LIBELLE') <> '') then
    begin
    if TobLigCom.detail.Count > 0 then
        begin
        for iNbRow := TobLigCom.Detail.Count - 1 downto 0 do
            begin
            if (TobLigCom.Detail [iNbRow].GetValue ('GCT_LIBELLE') <> '') and
               (TobLigCom.Detail [iNbRow].GetValue ('GCT_LIBELLE') <> ' ') then break;
            end;
        if DS.State <> dsInsert then TLBContenu.Clear;
        For iRow := 0 To iNbRow do
            begin
            TLBContenu.Add (TobLigCom.Detail[iRow].GetValue ('GCT_LIBELLE'));
            end;
        SetField ('GCT_CONTENU', TLBContenu.Text);
        end;
    end;
end;

function TOM_Commentaire.GetTobLigne (ARow : integer) : TOB;
begin
Result:=Nil ;
if (ARow < 0) or (ARow >= TobLigCom.Detail.Count) then Exit ;
Result := TobLigCom.Detail [ARow] ;
end;

procedure TOM_Commentaire.InitialiseGrille;
begin
GLigne.Enabled := True;
GLigne.VidePile(False) ;
GLigne.RowCount:= 50;
if TobLigCom <> Nil then
    begin
    TobLigCom.ClearDetail;
    end;
end;

procedure TOM_Commentaire.InitialiseLigne (ARow : integer) ;
Var TOBL : TOB ;
begin
TOBL := GetTOBLigne (ARow) ;
if TOBL <> Nil then TOBL.InitValeurs;
TOBL.PutValue ('GCT_LIBELLE', '');
GLigne.Rows [ARow].Clear;
end;

procedure TOM_Commentaire.RempliLesLignes;
var THDREOContenu : THRichEditOLE;
    iInd : integer;
begin
InitialiseGrille;
if DS.State <> dsInsert then
    begin
    THDREOContenu := THRichEditOLE (GetControl ('GCT_CONTENU'));
    for iInd := 0 to THDREOContenu.lines.Count - 1 do
        begin
        CreerTobLigne (iInd);
        TobLigCom.Detail[iInd].PutValue ('GCT_LIBELLE', THDREOContenu.lines.Strings[iInd]);
        end;
    TobLigCom.PutGridDetail (GLigne, False, False, 'GCT_LIBELLE', True);
    end;
GLigne.RowCount := GLigne.RowCount + 50;
stLibelle := GLigne.Cells[iColLibelle, 0];
end;

procedure TOM_Commentaire.SupprimeLigne (ARow : Longint) ;
begin
if ARow < 0 then Exit ;
if (ARow >= TobLigCom.Detail.Count) then Exit;

GLigne.CacheEdit;
GLigne.SynEnabled := False;
GLigne.DeleteRow (ARow);

TobLigCom.Detail.Delete(ARow);
GLigne.MontreEdit;
GLigne.SynEnabled := True;
if (DS.State <> dsEdit) and (DS.State <> dsInsert) then TFFicheListe(Ecran).Bouge(nbEdit);
SetField ('GCT_LIBELLE', GetField ('GCT_LIBELLE'));
end;

// Gesion des actions  *********************************************************

procedure TOM_Commentaire.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var Arow : integer;
    Cancel, Chg : boolean;
begin
if(Screen.ActiveControl = GLigne) then
    begin
    ARow := GLigne.Row;
    Case Key of
        VK_RETURN : Key:=VK_TAB ;
        VK_DELETE : begin
                    if Shift=[ssCtrl] then
                        begin
                        Key := 0 ;
                        SupprimeLigne (ARow) ;
                        Cancel := False;
                        Chg := False;
                        GLigneRowEnter (Sender, GLigne.Row, Cancel, Chg);
                        end ;
                    end;
        end;
    end;
end;

procedure TOM_Commentaire.GLigneExit(Sender: TObject);
var Cancel, Chg : boolean;
begin
if (GLigne.Cells[iColLibelle, GLigne.Row] <> stLibelle) then
   begin
   Cancel := False;
   Chg := False;
   GLigneRowExit (Sender, GLigne.Row, Cancel, Chg);
   stLibelle := GLigne.Cells[iColLibelle, GLigne.Row];
   end;
end;

procedure TOM_Commentaire.GLigneRowEnter (Sender: TObject; Ou: Integer;
                                          var Cancel: Boolean; Chg: Boolean);
var ARow : Integer;
    iRow : integer;
begin
bPasRowExit := False;
ARow := Ou;
if (ARow >= TobLigCom.detail.count) then
    begin
    iRow := TobLigCom.Detail.Count;
    while iRow <= Arow do
        begin
        CreerTobLigne (iRow);
        iRow := iRow + 1;
        end;
    Ou := ARow;
    GLigne.Row := Ou;
    end;
if Ou = GLigne.RowCount - 1 then
    begin
    GLigne.RowCount := GLigne.RowCount + 10;
    end;
stLibelle := GLigne.Cells [iColLibelle, Ou];
end;

procedure TOM_Commentaire.GLigneRowExit(Sender: TObject; Ou: Integer;
                                        var Cancel: Boolean; Chg: Boolean);
var Tobl : Tob;
begin
if (bPasRowExit) then exit;
Tobl := GetTobLigne (Ou);
if Tobl = nil then
    begin
    if TobLigCom.Detail.Count = 0 then
       begin
       stLibelle := GLigne.Cells[iColLibelle, Ou];
       CreerTobLigne (0);
       Tobl := GetTobLigne (0);
       GLigne.Cells[iColLibelle, Ou] := stLibelle;
       stLibelle := '';
       end;
    end;
if (Tobl <> nil) and (stLibelle <> GLigne.Cells[iColLibelle, Ou]) then
    begin
    TobL.PutValue ('GCT_LIBELLE', GLigne.Cells[iColLibelle, Ou]);
    if (DS.State <> dsEdit) and (DS.State <> dsInsert) then TFFicheListe(Ecran).Bouge(nbEdit);
    SetField ('GCT_LIBELLE', GetField ('GCT_LIBELLE'));
    end;
end;

// Procédures héritées *********************************************************

procedure TOM_Commentaire.OnNewRecord ;
begin
  inherited ;
if TobLigCom <> nil then TobLigCom.ClearDetail;
end ;

procedure TOM_Commentaire.OnDeleteRecord ;
begin
  inherited ;
if TobLigCom <> nil then TobLigCom.ClearDetail;
end ;

procedure TOM_Commentaire.OnUpdateRecord ;
begin
Enregistrer;
    inherited ;
end ;

procedure TOM_Commentaire.OnLoadRecord ;
begin
  inherited ;
if DS.State=dsInsert then SetField('GCT_TYPECOMMENT', 'LIG');
RempliLesLignes;
end ;

procedure TOM_Commentaire.OnArgument ( S: String );
begin
  inherited ;
TFFicheListe(Ecran).onKeyDown := FormKeyDown;

GLigne := THGrid(GetControl('GLIGNE'));
GLigne.OnRowEnter := GLigneRowEnter;
GLigne.OnRowExit := GLigneRowExit;
GLigne.OnExit := GLigneExit;

TobLigCom := Tob.Create ('', Nil, -1);
GLigne.ColLengths[iColLibelle] := 70;
end ;

procedure TOM_Commentaire.OnClose ;
begin
  inherited ;
TobLigCom.Free;
end ;

Initialization
  registerclasses ( [ TOM_Commentaire ] ) ;
end.
