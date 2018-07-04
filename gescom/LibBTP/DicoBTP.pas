unit DicoBTP;


interface
Uses classes,HEnt1,Paramsoc,stdctrls,ComCtrls,Controls,DBCtrls,sysutils,HCtrls,grids,dbgrids,menus,Ent1,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}hmsgbox,
{$IFDEF EAGLCLIENT}
     MenuOLX,
{$ELSE}
     MenuOLG,
{$ENDIF}
    HTB97,HDB,Forms,ExtCtrls ,checklst,Hout,UiUtil;

Const MaxDico = 30;

Var DicoAffaire : array[1..2,1..MaxDico] of string ;

// Fonctions de traduction vocab affaire / Gestion interne
function TraduitMotGA(mot : string; NumDico : integer =0) : string ;
function TraduitGA(phrase : string) : string ;
Procedure InitDicoAffaire ;
Procedure TraduitAFLibGridDB(Gr : TDBGrid );
Procedure TraduitAFLibGridST(Gr : TStringGrid );
Procedure TraduitTablette( NomEcran : String);
Procedure TraduitMenu(NumMenu : integer);
Procedure TraduireMetaDataAffaire; // traduction Dechamps + Detables

Procedure TraduitEcranAffaire (F:TForm);
procedure TraduitForm(Contain : TComponent);
//
Procedure PGIBoxAF (Msg, Titre : String; Prefixe : string = ''); overload;
procedure PGIBoxAF(Msg, Titre: string; const pArgs: array of const); overload;
//
procedure PGIInfoAF (Msg, Titre : String; Prefixe : string = ''); overload
procedure PGIInfoAF(Msg, Titre: string; const pArgs: array of const); overload;

function PGIAskCancelAF (Msg, Titre : string; Prefixe : string=''):integer ; overload;
function PGIAskCancelAF(Msg, Titre: string; const pArgs: array of const): integer; overload;

function PGIAskAF (Msg, Titre : string; Prefixe : string=''):integer ;
Function FindTitre(Titre, Prefixe : string):string;

// Fonction de gestion des niveaux de gestion d'affaires
Function VersionGAAvancee : Boolean;

implementation
uses EntGC
  ,CbpMCD
  ,CbpEnumerator
;

function TraduitMotGA(mot : string; NumDico : integer =0) : string ;
var i : integer ;
begin
result:=mot ;
If NumDico<>0 Then result:=DicoAffaire[2,NumDico]
Else               for i:=1 to MaxDico do if Uppercase(DicoAffaire[1,i])=Uppercase(mot) then result:=DicoAffaire[2,i] ;

if result<>mot then
   begin
    if (length(mot)>1) and (ord(mot[1])<=ord('Z')) and (ord(mot[2])>ord('Z')) then
       result:=FirstMajuscule(result)
    else if ord(mot[1])<=ord('Z') then result:=uppercase(result);
    End ;
end ;

function TraduitGA(phrase : string) : string ;
var mot,mot2,StFind : string ;
    i, ipos : integer;
begin
result := phrase;

// Recherche automatique des expressions du dico
for i:=1 to MaxDico do
    Begin
    StFind := Uppercase(DicoAffaire[1,i]);
    ipos:=Pos(StFind,Uppercase(Phrase));
    While ipos>0 do
        BEGIN
        mot := Copy(phrase, ipos, Length(StFind));
        mot2:=TraduitMotGA(mot,i) ;
        Phrase:=FindEtReplace(phrase,mot,mot2,true) ;
        ipos:=Pos(StFind,Uppercase(phrase));
        End ;
    End;
Result := Phrase;
end ;

Procedure TraduitAFLibGridDB(Gr : TDBGrid );
Var i : integer;

Begin
{$IFDEF EAGLCLIENT}
//AFAIREEAGL
{$ELSE}
If Not(Gr is TDBGrid) Then Exit;
if (Gr.Columns.Count = 1) then Exit ;
For i:=0 to Gr.Columns.Count-1 do
    Begin
    if Gr.columns[i].Field <> Nil then
        Gr.columns[i].Field.DisplayLabel:=TraduitGA(Gr.columns[i].Field.DisplayLabel);
    End;
{$ENDIF}
End;

Procedure TraduitAFLibGridST(Gr : TStringGrid );
Var i : integer;

Begin
{$IFDEF EAGLCLIENT}
//AFAIREEAGL
{$ELSE}
If Not(Gr is TStringGrid) Then Exit;
if (Gr.ColCount = 1) then Exit ;
if (Gr.FixedRows = 0) then Exit ;
For i:=0 to Gr.ColCount-1 do
    Begin
    Gr.Cells[i, 0]:=TraduitGA(Gr.Cells[i, 0]);
    End;
{$ENDIF}
End;

Procedure InitDicoAffaire ;
Var i : integer;
Begin
i := 1;
if ctxSCOT in V_PGI.PGIContexte Then
    Begin
    // Attention l'ordre a de l'importance pour déterminer les expressions à traduire en priorité ...
    DicoAffaire[1,i] := 'gestion d''affaires';  DicoAffaire[2,i] := 'gestion de chantiers';  Inc(i);
    DicoAffaire[1,i] := 'l''affaire';           DicoAffaire[2,i] := 'le chantier' ;      Inc(i);
    DicoAffaire[1,i] := 'd''affaire';           DicoAffaire[2,i] := 'de chantier' ;      Inc(i);
    DicoAffaire[1,i] := 'affaires';             DicoAffaire[2,i] := 'chantiers';         Inc(i);
    DicoAffaire[1,i] := 'affaire';              DicoAffaire[2,i] := 'chantier' ;         Inc(i);
    DicoAffaire[1,i] := 'la ressource';         DicoAffaire[2,i] := 'l''intervenant' ;    Inc(i);
    DicoAffaire[1,i] := 'code ressource';       DicoAffaire[2,i] := 'code intervenant' ;  Inc(i);
    DicoAffaire[1,i] := 'une ressource';        DicoAffaire[2,i] := 'un intervenant' ;    Inc(i);
    DicoAffaire[1,i] := 'de ressource';         DicoAffaire[2,i] := 'd''intervenant' ;    Inc(i);
(*10*)    DicoAffaire[1,i] := 'ressources';           DicoAffaire[2,i] := 'intervenants' ;      Inc(i);
    DicoAffaire[1,i] := 'ressource';            DicoAffaire[2,i] := 'intervenant';        Inc(i);
    DicoAffaire[1,i] := 'tempo';                DicoAffaire[2,i] := 'BTP';             Inc(i);
    DicoAffaire[1,i] := 'employés';             DicoAffaire[2,i] := 'intervenants' ;      Inc(i);
    DicoAffaire[1,i] := 'employé';              DicoAffaire[2,i] := 'intervenant';        Inc(i);
    DicoAffaire[1,i] := 'l''article';           DicoAffaire[2,i] := 'la fourniture';    Inc(i);
    DicoAffaire[1,i] := 'd''articles';           DicoAffaire[2,i] :='de fournitures';  Inc(i);
    DicoAffaire[1,i] := 'un article';           DicoAffaire[2,i] := 'une fourniture';   Inc(i);
    DicoAffaire[1,i] := 'articles';             DicoAffaire[2,i] := 'fournitures';      Inc(i);
    DicoAffaire[1,i] := 'article';              DicoAffaire[2,i] := 'fourniture';
(*20*) 

    // Attention limité à 30 ...
    End
Else
    Begin
    DicoAffaire[1,i] := 'Articles et stocks';              DicoAffaire[2,i] := 'Stocks';
    End;
End;

Procedure PGIBoxAF (Msg, Titre : String; Prefixe : string = '');
Begin
Titre := FindTitre(Titre,Prefixe);
PGIBox (TraduitGA(Msg),TraduitGA(titre));
End;

procedure PGIBoxAF(Msg, Titre: string; const pArgs: array of const);
begin
  if (Titre = '') then
    Titre := TitreHalley;
  PGIBox(Format(TraduitGA(Msg), pArgs), TraduitGA(titre));
end;

procedure PGIInfoAF (Msg, Titre : string; Prefixe : string='');
Begin
Titre := FindTitre(Titre,Prefixe);
PGIInfo (TraduitGA(Msg),TraduitGA(titre));
End;

procedure PGIInfoAF(Msg, Titre: string; const pArgs: array of const);
begin
  if (Titre = '') then
    Titre := TitreHalley;
  PGIInfo(Format(TraduitGA(Msg), pArgs), TraduitGA(titre));
end;

function PGIAskCancelAF (Msg, Titre : string; Prefixe : string='') :integer ;
Begin
Titre := FindTitre(Titre,Prefixe);
Result := PGIAskCancel(TraduitGA(Msg),TraduitGA(titre));
End;

function PGIAskCancelAF(Msg, Titre: string; const pArgs: array of const): integer;
begin
  if (Titre = '') then
    Titre := TitreHalley;
  Result := PGIAskCancel(Format(TraduitGA(Msg), pArgs), TraduitGA(titre));
end;

function PGIAskAF (Msg, Titre : string; Prefixe : string=''):integer ;
Begin
Titre := FindTitre(Titre,Prefixe);
Result := PGIAsk(TraduitGA(Msg),TraduitGA(titre));
End;

Function FindTitre(Titre, Prefixe : string):string;
Begin
{***** en cours de modif GMGMG
stTitre := '';

If Prefixe <> '' Then
    Begin
    stTable := PrefixeToTable(Prefixe);
    If (stTable <> '') Then stTitre := TableToLibelle (stTable);
    Exit;
    End
Else stTitre := Titre;
Result := TraduitGA(stTitre);
}
 result := titre;
If (Result ='') Then Result := TitreHalley;
End;

(******************************************************************************)
// TraduitForm : procédure de traduction automatique de tous les textes des composants
// contenus dans le component passé en paramètre et ses components enfants
//
// Appel courant : dans le FormShow d'un TForm : TraduitForm(Self)
//
// entrée : TComponent
//
// Remarque : la procédure est récursive pour les components pouvant contenir
// d'autres composants (TForm, TPanel, TTabSheet, TPageControl, TgroupBox)
//
// Auteur : Marie-Christine Desseignet le 25/01/2000
(******************************************************************************)
procedure TraduitForm(Contain : TComponent);
var
i,j,value:integer;

begin
// TraduitTablette('');
//
// Traduire les conteneurs qui ont des captions à traduire avant la traduction de leur contenu
//
If (Contain Is TForm) Then TForm(Contain).caption :=TraduitGA(TForm(Contain).caption)
Else
If (Contain Is TGroupBox) Then TForm(Contain).caption :=TraduitGA(TGroupBox(Contain).caption);

//
// Boucle sur les composants contenus dans le component courant
//
for i:=0 to Contain.ComponentCount-1 do
begin
// Si le component courant est du type TControl
if (Contain.components[i] is TControl) then
   begin
    If (Contain.Components[i] is TButtonControl) then
        begin
        TButtonControl(Contain.Components[i]).Hint := TraduitGA(TButtonControl(Contain.Components[i]).Hint);
        if (Contain.Components[i] is TCheckBox) then
            TCheckBox(Contain.Components[i]).caption :=TraduitGA(TCheckBox(Contain.Components[i]).caption);
        if (Contain.Components[i] is TDBCheckBox) then
            TCheckBox(Contain.Components[i]).caption :=TraduitGA(TCheckBox(Contain.Components[i]).caption);
        end
    Else
    If (Contain.Components[i] is TCustomEdit) then
        begin
        TCustomEdit(Contain.Components[i]).Hint := TraduitGA(TCustomEdit(Contain.Components[i]).Hint);
        end
    Else
    If (Contain.Components[i] is TLabel) then
        begin
        TLabel(Contain.Components[i]).caption :=TraduitGA(TLabel(Contain.Components[i]).caption);
        TLabel(Contain.Components[i]).Hint := TraduitGA(TLabel(Contain.Components[i]).Hint);
        end
    Else
    If (Contain.Components[i] is TStaticText) then
        begin
        TStaticText(Contain.Components[i]).caption :=TraduitGA(TStaticText(Contain.Components[i]).caption);
        TStaticText(Contain.Components[i]).Hint :=TraduitGA(TStaticText(Contain.Components[i]).Hint);
        end
    Else
    If (Contain.Components[i] is TCheckListBox) then
        begin
        TChecklistBox(Contain.Components[i]).items.Commatext :=TraduitGA(TChecklistBox(Contain.Components[i]).items.Commatext);
        TChecklistBox(Contain.Components[i]).Hint :=TraduitGA(TChecklistBox(Contain.Components[i]).Hint);
        end
    Else
    if (Contain.Components[i] is TListBox) then
        begin
        TlistBox(Contain.Components[i]).items.Commatext :=TraduitGA(TlistBox(Contain.Components[i]).items.Commatext);
        TlistBox(Contain.Components[i]).Hint :=TraduitGA(TlistBox(Contain.Components[i]).Hint);
        end
    Else
    if (Contain.Components[i] is TComboBox) then
        Begin
        TComboBox(Contain.Components[i]).Hint := TraduitGA(TComboBox(Contain.Components[i]).Hint);
        Value := TComboBox(Contain.Components[i]).itemindex;
        TComboBox(Contain.Components[i]).items.Commatext :=TraduitGA(TComboBox(Contain.Components[i]).items.Commatext);
        TComboBox(Contain.Components[i]).itemindex := Value;
        End
    Else
    If (Contain.Components[i] is TGraphicControl) then
        begin
        TGraphicControl(Contain.Components[i]).Hint := TraduitGA(TGraphicControl(Contain.Components[i]).Hint);
        end
    Else
    if (Contain.Components[i] is TStringGrid) then
        TraduitAFLibGridST(TStringGrid(Contain.Components[i]))
    Else
    if (Contain.Components[i] is TDBGrid) then
        TraduitAFLibGridDB(TDBGrid(Contain.Components[i]))
    Else
   // Appel récursif de la fonction de traduction pour les containers
   If (Contain.Components[i] Is TForm) or (Contain.Components[i] Is TTabSheet) or
      (Contain.Components[i] Is TPageControl) or (Contain.Components[i] Is TgroupBox) or
      (Contain.Components[i] Is Tpanel) Then TraduitForm(Contain.Components[i]);
   end
else
// Si le component courant est du type TMenu
if (Contain.components[i] is TMenu) then
   begin
   if (Contain.components[i] is Tpopupmenu) then
      for j:=0 to Tpopupmenu(Contain.components[i]).Items.Count-1 do
         TMenuItem(Tpopupmenu(Contain.components[i]).items[j]).Caption :=TraduitGA(TMenuItem(Tpopupmenu(Contain.components[i]).items[j]).Caption);
   end
else
// Si le component courant est du type THMsgBox
if (Contain.components[i] is THMsgBox) then
   begin
   if (Contain.components[i] is THMsgBox) then
      THMsgBox(Contain.components[i]).Mess.Commatext :=TraduitGA(THMsgBox(Contain.components[i]).Mess.Commatext)
   end;
end;
end;

Procedure TraduitEcranAffaire (F:TForm);
BEGIN
TraduitForm(F);
END;

Function ExtractInfo( St : String ; indice : integer ) : String ;
Var Pos1 : Longint ;
    i    : integer ;
BEGIN
Result:='' ;
for i:=0 to Indice do
    BEGIN
    Pos1:=Pos(#9,St) ;
    if Pos1>1 then BEGIN Result:=Copy(St,1,Pos1-1) ; Delete(St,1,Pos1) ; END
              else Result:='' ;
    END ;
END ;

Procedure TraduitTablette( NomEcran : String);
Var i,j : integer;
    Stt : HTstrings;
    Tablette, Lib,Fin,Debut : string;
BEGIN
Tablette:='AFTTYPEPREVU' ; i:=TTToNum(Tablette) ;
if i<0 then exit;
Stt:=V_PGI.DECombos[i].Valeurs ;
//if Stt=Nil then RemplirListe(Tablette,'') ;
//Stt:=V_PGI.DECombos[i].Valeurs ;
if Stt=Nil then exit;
for j:=0 to Stt.Count-1 do
    BEGIN
    //St := Stt[j];
    (*(Pos1:=Pos(#9,St) ; Debut := copy (St,1,Pos1-1);
    Delete(St,1,Pos1) ; Pos2:= Pos(#9,St);
    Lib:= Copy(St,1,Pos2-1);
    Fin:= Copy (St,Pos2+1,Length(St)-Pos2-1); **)
    Debut := ExtractInfo(Stt[j],0);
    Lib   := ExtractInfo(Stt[j],1);  Lib := traduitGA(Lib);
    fin   := ExtractInfo(Stt[j],2);
    Stt[j] := Debut+'#9'+Lib+'#9'+Fin;
    END;
// Lookup
//V_PGI.DECombos[ii].Libelle
END;

Function GetLibItem(nTag : integer) : String ;
var TreeN : TTreeNode ;
BEGIN
result := '';
{$IFDEF EAGLCLIENT}
//AFAIREEAGL
{$ELSE}
//mItem:=THOutItem(GetTagItem(FMenuG.OutLook,nTag)) ;
//if mItem <> nil then Result := mItem.Titre;
TreeN:=GetTreeItem(FMenuG.Treeview,nTag) ;
if TreeN<>Nil then result := TreeN.Text ;
{$ENDIF}
END;

Procedure TraduitMenu(NumMenu : integer);
Var ListeItems, Item, Lib : String;
    NumItem : integer;
    Bt : TToolbarbutton97;
    mn : TMenuItem ;
BEGIN
Listeitems := '';
Case NumMenu of
    32 :
        BEGIN
        FMenuG.PTitre.Caption := 'Stocks';
        FMenuG.RemoveGroup(32500,True);
        FMenuG.RemoveGroup(32700,True);
        END;
    71 :
         BEGIN
         ListeItems := '71102;71103;71104;-71220;71221;71222;71223;-71230;71231;71232;71233;71234';
         ListeItems := ListeItems + ';71205;71301;71401;71513';
         END;
    72 : ListeItems :='-72101;72104;-72110;72113;-72220;-72230;72221;72222;72231;72232;72211;72212';
    73 : ListeItems := '73402';
    74 : ListeItems := '-74330;-74350;-74460;-74430;74602;74251;74252';
    END;
Item:=(Trim(ReadTokenSt(ListeItems)));
While (Item <>'') do
    BEGIN
    NumItem := strToInt(Item);
    Lib := TraduitGA(GetLibItem(NumItem));
    FMenuG.RenameItem(NumItem,Lib) ;
    Item:=(Trim(ReadTokenSt(ListeItems)));
    END;

if (numMenu = 71) then
    BEGIN // Modification du module de base
    if (ctxScot In V_PGI.PGIContexte) then
        BEGIN
        Bt := TToolBarButton97(FMenuG.FindComponent ( 'bModule' + IntToStr(NumMenu)));
        if Bt <> nil then Bt.Caption := 'Gestion Interne';
        mn := FindMenuItem(NumMenu,FmenuG.mnModules);
        if mn <> Nil then mn.Caption := 'Gestion Interne';
        FMenuG.PTitre.Caption := 'Gestion Interne';
        FMenuG.RemoveGroup(71600,True); // Suppression des tarifs
        END
    else
        BEGIN
        FMenuG.RemoveItem(71302); // suppression edition lettres de mission
        END;
    END else
if (numMenu = 74) then
    BEGIN
{$ifndef DEBUG}
    FMenuG.RemoveItem(74610); FMenuG.RemoveItem(74611); // Test en debug
{$ENDIF}
    if (ctxScot In V_PGI.PGIContexte) then
        BEGIN
        FMenuG.RemoveItem(74315);   // Tarif tiers
        FMenuG.RemoveItem(74346);   // Tarif article
        FMenuG.RemoveItem(74351);   // Tarif ressource
        END;
    END;
END;

Procedure TraduireUneTable(Prefixe : string);
Var
  Mcd : IMCDServiceCOM;
  Table     : ITableCOM ;
  FieldList : IEnumerator ;
	Field     : IFieldCOMEx ;
BEGIN

  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  Table := Mcd.GetTable(mcd.PrefixeToTable(prefixe));
  FieldList := Table.Fields;
  FieldList.Reset();
  While FieldList.MoveNext do
  begin
  	Field := FieldList.Current as IFieldComex;
    if Field.name <>'' then
       BEGIN
       	Field.Libelle:=TraduitGA(Field.Libelle);
       END;
  end;
END;

Procedure TraduireMetaDataAffaire;
BEGIN
if not(ctxAffaire in V_PGI.PGIContexte) then Exit;
If Not(ctxScot in V_PGI.PGIContexte) Then  Exit;
TraduireUneTable ('AFF');
TraduireUneTable ('AFA');
TraduireUneTable ('ARS');
TraduireUneTable ('T');
END;

Function VersionGAAvancee : Boolean;
BEGIN
Result := true;
if (ctxAffaire in V_PGI.PGIContexte) and (ctxNegoce in V_PGI.PGIContexte ) then Result := True;
// A voir en fonction des seria mises en place par la suite
END;
end.

