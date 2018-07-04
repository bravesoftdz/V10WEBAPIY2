{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 04/04/2002
Modifié le ... :   /  /
Description .. : Rectification des tables libres
Mots clefs ... : LIBRES;PAIE
*****************************************************************
Structure :
Ce source repose sur une tof Mere TOF_PGREAFFECT qui gère l'évèmentiel de la grille nommé GRILLE,
l'affichage des champs et valeurs de prefixe PPU
Les tofs fille restent personnalisés au traitement spécifique de chaque réaffectation (fiche)
}
unit UTOFPGReaffectTable;

interface
uses StdCtrls,Controls,Classes,Graphics,forms,sysutils,ComCtrls,ParamSoc,
{$IFDEF EAGLCLIENT}
     eQRS1,
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     HCtrls,HEnt1,HMsgBox,UTOF,PGIEnv,EntPaie,UTOB,Pgoutils,ed_tools,HStatus,
     HQRy,HTB97,P5Util,Grids,HSysMenu,PGEditOutils,PGEdtEtat;

Type

    TOF_PGREAFFECT= Class (TOF)
       procedure OnArgument(Arguments : String ) ; override ;
       procedure OnClose;  Override;
       private
       Grille : THGrid;
       Tob_Grille : Tob;
       Modifier,ZoneAModifier : Boolean ;
       TitreColonne : TStringList;
       procedure ExitEdit(Sender: TObject);
       procedure GrillePostDrawCell(ACol,ARow : Longint; Canvas : TCanvas ; AState: TGridDrawState);
       procedure GrilleCellEnter(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
       procedure GrilleCellExit(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
       procedure GridSelectCell(Sender: TObject; ACol, ARow: Longint; var CanSelect: Boolean) ;
       procedure GridColorCell (ACol,ARow : Longint; Canvas : TCanvas ; AState: TGridDrawState);
       Procedure RecupClauseChampReaffecter(IndiceChamp,Row : integer; var st: string);
     END ;

     TOF_PGREAFFECTTABLE= Class (TOF_PGREAFFECT)
       procedure OnArgument(Arguments : String ) ; override ;
       procedure OnUpdate; Override;
       private
       GlbSql,Pref : string;
       procedure MiseEnFormeGrille;
       Procedure MiseEnFormeColGrille(Champ,Suffixe : string);
       procedure OnChangeListeTable(Sender : TObject);
       Procedure OnChangeListeChamp(Sender : TObject);
       Procedure OnClickBtnActiveWhere(Sender : TObject);
       Procedure OnClickBtnMajValue(Sender : TObject);
       Function  RecupTabletteAssocie (Champ : string) : String;
       Function  RecupNomTableAssocie (Champ : string) : String;
       procedure RecopieTablesSalaries (var NbMajOk : integer);
       procedure MiseAjourSelectif(var NbMajOk : integer);
       Function  RendClauseAbsenceSalarie ( StClause : String) : String;
     END ;

     TOF_PGREAFFECTORGANISME= Class (TOF_PGREAFFECT)
       procedure OnArgument(Arguments : String ) ; override ;
       procedure OnUpdate; Override;
       Private
       procedure MiseEnFormeGrille;
       Procedure OnClickBtnActiveWhere(Sender : TObject);
       Procedure OnClickBtnMajValue(Sender : TObject);
       Procedure OnClickHisto(Sender : TObject);
     End;

const
	// libellés des messages
	TexteMessage: array[1..12] of string 	= (
          {1}        'Veuillez saisir un domaine de réaffectation.',
          {2}        'Attention,la réaffectation de certaines données seront perdues!#13#10Voulez-vous réactualiser la grille?',
          {3}        'Veuillez saisir une valeur à réaffecter.' ,
          {4}        'La zone de réaffectation est inconnue.',
          {5}        'La zone de réaffectation n''est pas intégrée à la grille.',
          {6}        'Veuillez selectionner un champ libre à réaffecter.',
          {7}        'Aucunes réaffectations n''ont été apportés aux données.',
          {8}        'Aucunes réaffectations n''ont été sauvegardés.',
          {9}        'Attention!Vous demandez à recopier les champs libres de la table salarié#13#10 sur les champs libres de l''historique de paie.#13#10 Voulez-vous continuer le traitement?',
          {10}       'Attention!Vous demandez à mettre à jour les champs libres de l''historique de paie.#13#10 Voulez-vous continuer le traitement?',
          {11}       'Veuillez cliquer sur le bouton de validation vert pour la recopie des données.',
          {12}       ''
                    );

implementation

{ TOF_PGREAFFECT }

procedure TOF_PGREAFFECT.ExitEdit(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then
  if edit.text<>'' then
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end;


procedure TOF_PGREAFFECT.GridColorCell(ACol, ARow: Integer;
  Canvas: TCanvas; AState: TGridDrawState);
begin
if not ZoneAModifier then exit;
if Pos('modifié',Grille.cells[ACol,0])>=1 then
   If Grille.Cells[Acol,Arow]<>Grille.Cells[Acol-1,Arow] then
     if (Grille.Cells[Acol,Arow]<>'') AND (Grille.CellValues[Acol,Arow]<>'NON') then
       Canvas.Font.Color := clRed ;

end;

procedure TOF_PGREAFFECT.GridSelectCell(Sender: TObject; ACol,
  ARow: Integer; var CanSelect: Boolean);
begin
{if not ZoneAModifier then exit;
CanSelect := (Pos('modifié',Grille.cells[Acol,0])>0);
if (Not CanSelect) AND (Grille.ColCount>2)  then
  Begin
  If (Acol = Grille.ColCount-1) then
     begin
     // En fin de ligne on se positionne au début de la ligne suivante
     Grille.col := 0;
     if (ARow <> Grille.RowCount-1) then
        Grille.row := ARow + 1;
     end
  else
     Grille.col := Acol+1;
  End;}
end;

procedure TOF_PGREAFFECT.GrilleCellEnter(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
begin
if not ZoneAModifier then exit;
if Pos('modifié',Grille.cells[Grille.col,0])<1 then
     begin
     Cancel:=True;
     If (Acol = Grille.ColCount-1) then
       begin
       // En fin de ligne on se positionne au début de la ligne suivante
       Grille.col := 0;
       if (ARow <> Grille.RowCount-1) then
          Grille.Row := Grille.Row + 1;
       end
     else
       Grille.Col := Grille.Col+1;
     end;
end;

procedure TOF_PGREAFFECT.GrilleCellExit(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
begin
if not ZoneAModifier then exit;
{if (Pos('modifié',Grille.cells[ACol,0])>=1) And (Grille.Row=Arow)   then
   If (Acol = Grille.ColCount-1) then
     begin
     // En fin de ligne on se positionne au début de la ligne suivante
     if (Arow <> Grille.RowCount-1) then
        Grille.row := Arow + 1
     else
        Grille.row := 1;
     Grille.col := 3;
     end
   else
     Grille.col := Acol+2;  }
end;

procedure TOF_PGREAFFECT.GrillePostDrawCell(ACol, ARow: Integer;
  Canvas: TCanvas; AState: TGridDrawState);
begin
if not ZoneAModifier then exit;
if Pos('modifié',Grille.Cells[ACol,0])<1 then
  GridGriseCell(Grille,Acol,Arow,Canvas)
else
   if Acol>0 then 
   If Grille.Cells[Acol,Arow]<>Grille.Cells[Acol-1,Arow] then
     if (Grille.Cells[Acol,Arow]<>'') AND (Grille.CellValues[Acol,Arow]<>'NON') then
       Canvas.Pen.Color := clRed ;
end;

procedure TOF_PGREAFFECT.OnArgument(Arguments: String);
Var
Min,Max : string;
Edit: ThEdit;
begin
  inherited;
Modifier:=False;
ZoneAModifier:=False;
VisibiliteChamp (Ecran);
VisibiliteChampLibre (Ecran);
//Valeur par défaut
RecupMinMaxTablette('PG','SALARIES','PSA_SALARIE',Min,Max);
Edit:=ThEdit(getcontrol('PPU_SALARIE'));
If Edit<>nil then Begin Edit.text:=Min; Edit.OnExit:=ExitEdit; End;
Edit:=ThEdit(getcontrol('PPU_SALARIE_'));
If Edit<>nil then Begin Edit.text:=Max; Edit.OnExit:=ExitEdit;End;
RecupMinMaxTablette('PG','ETABLISS','ET_ETABLISSEMENT',Min,Max);
Edit:=ThEdit(getcontrol('PPU_ETABLISSEMENT'));
If Edit<>nil then Begin Edit.text:=Min;  End;
Edit:=ThEdit(getcontrol('PPU_ETABLISSEMENT_'));
If Edit<>nil then Begin Edit.text:=Max; End;
Grille := THGrid(GetControl('GRILLE'));
if Grille<>nil then
  Begin
  Grille.PostDrawCell := GrillePostDrawCell;
//  Grille.OnCellEnter := GrilleCellEnter;

//  Grille.GetCellCanvas := GridColorCell ;
//  Grille.OnCellExit:=GrilleCellExit;
 {
  Grille.OnSelectCell:=GridSelectCell; }
  End;
TitreColonne:=TStringList.Create;

SetControlVisible('CKTRAVAILN1',(VH_PAIE.PGLibelleOrgStat1<>''));
SetControlVisible('CKTRAVAILN2',(VH_PAIE.PGLibelleOrgStat2<>''));
SetControlVisible('CKTRAVAILN3',(VH_PAIE.PGLibelleOrgStat3<>''));
SetControlVisible('CKTRAVAILN4',(VH_PAIE.PGLibelleOrgStat4<>''));
SetControlVisible('CKCODESTAT',(VH_PAIE.PGLibCodeStat<>''));
SetControlVisible('CKLIBREPCMB1',(VH_PAIE.PgLibCombo1<>''));
SetControlVisible('CKLIBREPCMB2',(VH_PAIE.PgLibCombo2<>''));
SetControlVisible('CKLIBREPCMB3',(VH_PAIE.PgLibCombo3<>''));
SetControlVisible('CKLIBREPCMB4',(VH_PAIE.PgLibCombo4<>''));
end;



procedure TOF_PGREAFFECT.OnClose;
begin
  inherited;
if TitreColonne<>nil then TitreColonne.Free;
if Tob_Grille<>nil then Tob_Grille.free;
end;

procedure TOF_PGREAFFECT.RecupClauseChampReaffecter(IndiceChamp,Row : integer; var st: string);
begin
St:='';
Repeat        // Recherche par Colonnes des valeurs modifiées
  if (Grille.CellValues[IndiceChamp,Row])<>(Grille.CellValues[IndiceChamp+1,Row]) then
    Begin
    if St<>'' then St:=St+',';
    St:=St+TitreColonne.Strings[IndiceChamp]+'="'+Grille.CellValues[IndiceChamp+1,Row]+'" ';
    End;
  IndiceChamp:=IndiceChamp+2;
Until IndiceChamp>=Grille.ColCount-1;
end;

{ TOF_PGREAFFECTTABLE }



procedure TOF_PGREAFFECTTABLE.OnArgument(Arguments: String);
var
Combo : THValComboBox;
Btn : TToolBarButton97 ;
begin
  inherited;
//Gestionnnaire d'evenement
Combo := THValComboBox(GetControl('LISTETABLE'));
if combo<>nil then combo.OnChange:=OnChangeListeTable;
Combo := THValComboBox(GetControl('LISTECHAMP'));
if Combo<>nil then   Combo.OnChange:=OnChangeListeChamp;
Btn := TToolBarButton97(GetControl('BTNACTIVEWHERE'));
if btn<>nil then Btn.Onclick:=OnClickBtnActiveWhere;
Btn := TToolBarButton97(GetControl('BTNMAJVALUE'));
if btn<>nil then Btn.Onclick:=OnClickBtnMajValue;
end;



procedure TOF_PGREAFFECTTABLE.OnChangeListeTable(Sender: TObject);
begin
SetControlVisible('PANEL',(GetControlText('LISTETABLE')<>'RECOPIE'));
SetControlVisible('DATEDEB',(GetControlText('LISTETABLE')<>'SALARIES'));
SetControlVisible('DATEFIN',(GetControlText('LISTETABLE')<>'SALARIES'));
SetControlVisible('TDATEDEB',(GetControlText('LISTETABLE')<>'SALARIES'));
SetControlVisible('TDATEFIN',(GetControlText('LISTETABLE')<>'SALARIES'));
SetControlEnabled('BValider',(GetControlText('LISTETABLE')='RECOPIE'));
end;

procedure TOF_PGREAFFECTTABLE.OnChangeListeChamp(Sender: TObject);
var champ : string;
begin
champ:=GetControlText('LISTECHAMP');
SetControlProperty('OLDVALUE','DataType','');
SetControlProperty('NEWVALUE','DataType','');
If champ<>'' then
  Begin
  Champ:=RecupTabletteAssocie (Champ);
  SetControlProperty('NEWVALUE','DataType',Champ);
  SetControlProperty('OLDVALUE','DataType',Champ);
  SetControlText('NEWVALUE','');
  SetControlText('OLDVALUE','');
  End;
end;


procedure TOF_PGREAFFECTTABLE.OnClickBtnActiveWhere(Sender: TObject);
Var ListeTable,StSql,ChSql,Critere,AutrePref : string;
Q : Tquery;
Pages:TPageControl;
i,Col,Row,IndiceChamp : integer;
T:Tob;
HMTrad: THSystemMenu;
begin
Pref:='';
ListeTable:=GetControlText('LISTETABLE');
if (ListeTable='') then begin PGIBox(TexteMessage[1],Ecran.caption); exit; end;
if (ListeTable='RECOPIE') then begin PGIBox(TexteMessage[11],Ecran.caption); exit; end;
if Modifier then
 if PGIAsk(TexteMessage[2],Ecran.caption)=MrNo then
    exit;
Pref:=TableToPrefixe(ListeTable);
if Pref='PSA' then AutrePref:='PPU' else AutrePref:='PSA';
MiseEnFormeGrille;
Pages:=TPageControl(GetControl('Pages'));
if Pages<>nil then Critere:=RecupWhereCritere(Pages)
else Critere:='';
if (GlbSql<>'') then GlbSql:=ConvertPrefixe(GlbSql,AutrePref,Pref);
if Critere<>'' then
  Begin
  Critere:=ConvertPrefixe(Critere,AutrePref,Pref);
  if Pref='PPU' then Critere:=Critere+'AND PPU_DATEFIN>="'+USDateTime(StrToDate(GetControlText('DATEDEB')))+'" '+
  'AND PPU_DATEFIN<="'+USDateTime(StrToDate(GetControlText('DATEFIN')))+'"';
  End
else
  if Pref='PPU' then Critere:='WHERE PPU_DATEFIN>="'+USDateTime(StrToDate(GetControlText('DATEDEB')))+'" '+
  'AND PPU_DATEFIN<="'+USDateTime(StrToDate(GetControlText('DATEFIN')))+'"';
if GlbSql<>'' then
  ChSql:=Pref+'_ETABLISSEMENT,'+Pref+'_SALARIE,'+GlbSql
else
  ChSql:=Pref+'_ETABLISSEMENT,'+Pref+'_SALARIE ';
StSql:='SELECT DISTINCT '+ChSql+' FROM '+ListeTable+' '+Critere+' ORDER BY '+PRef+'_SALARIE' ;
Q:=OpenSql(StSql,True);
if Tob_Grille<>nil then begin Tob_Grille.free; Tob_Grille:=nil; end;
Tob_Grille:=Tob.create('La liste à maj',nil,-1);
Tob_Grille.LoadDetailDB('La liste à maj','','',Q,False);
Ferme(Q);

Row:=0;
Grille.RowCount:=Tob_Grille.Detail.Count+1;
For I:=0 to Tob_Grille.Detail.Count-1 do
  Begin
  T:=Tob_Grille.Detail[i];
  Row:=Row+1;
  IndiceChamp:=999;
  With T do
    For Col:=0 to Grille.ColCount-1 do
      Begin
      if Pos('modifié',Grille.Cells[Col,0])<1 then
        Begin
        IndiceChamp:=IndiceChamp+1;
        if GetNomChamp(IndiceChamp)<>'' then
          Begin
          if Pos('affecté',Grille.Cells[Col,0])>=1 then
            Begin
             Critere:=Grille.ColFormats[Col+1];
             ReadTokenPipe(Critere,'=');
             Grille.Cells[Col,Row]:=RechDom(Critere,GetValue(GetNomChamp(IndiceChamp)),False);
             Grille.Cells[Col+1,Row]:=RechDom(Critere,GetValue(GetNomChamp(IndiceChamp)),False);
            End
          else
            Grille.Cells[Col,Row]:=GetValue(GetNomChamp(IndiceChamp));
          End;
        End;
      End;
  End;
HMTrad.ResizeGridColumns(Grille);
if Grille.ColCount>2 then
  if Pref='PSA' then Grille.col:=3
    else
      if Pref='PPU' then Grille.col:=4; 
Grille.Enabled:=ZoneAModifier;
SetControlEnabled('BTNMAJVALUE',True);
SetControlEnabled('PNMODIF',True);
SetControlEnabled('BValider',True);
end;
//Affectation de la zone de modication
procedure TOF_PGREAFFECTTABLE.OnClickBtnMajValue(Sender: TObject);
var champ,Libelle,OldValue,NewValue,num : string;
Col,Row,IndiceChamp,Compteur : integer;
Obligatoire : Boolean;
begin
if Grille=nil then exit;
champ := GetControlText('LISTECHAMP');
OldValue:= GetControlText('OLDVALUE');
NewValue:= GetControlText('NEWVALUE');
Libelle:=''; IndiceChamp:=-1; Compteur:=0;
If (champ<>'') then
  Begin
  Obligatoire:=True;
  if champ='SO_PGLIBCODESTAT' then Libelle:=VH_Paie.PGLibCodeStat
  else
    if Pos('SO_PGLIBORGSTAT',champ)>0 then
       begin
       num:=Copy(Champ,Length(Champ),1);
       if num='1' then Libelle:=VH_Paie.PGLibelleOrgStat1
         else if num='2' then Libelle:=VH_Paie.PGLibelleOrgStat2
           else if num='3' then Libelle:=VH_Paie.PGLibelleOrgStat3
             else if num='4' then Libelle:=VH_Paie.PGLibelleOrgStat4;
       End
    else
      if Pos('SO_PGLIBCOMBO',champ)>0 then
         begin
         Obligatoire:=False;
         num:=Copy(Champ,Length(Champ),1);
         if num='1' then Libelle:=VH_Paie.PgLibCombo1
           else if num='2' then Libelle:=VH_Paie.PgLibCombo2
             else if num='3' then Libelle:=VH_Paie.PgLibCombo3
               else if num='4' then Libelle:=VH_Paie.PgLibCombo4;
         End;
  if (Obligatoire) and ((OldValue='') or (NewValue='')) then Begin PgiBox(TexteMessage[3],Ecran.caption); exit; End;
  if Libelle='' then begin PgiBox(TexteMessage[4],Ecran.caption); exit; end;
  For  Col:=0 to Grille.ColCount-1 do
    if Pos(Libelle ,Grille.Cells[Col,0])>0 then
       begin
       IndiceChamp:=Col;
       Break;
       End;
  if IndiceChamp=-1 then begin PgiBox(TexteMessage[5],Ecran.caption); exit; end;
  For Row:=1 to Grille.RowCount-1 do
    Begin
    if (Grille.CellValues[IndiceChamp,Row]=OldValue) and (Grille.CellValues[IndiceChamp+1,Row]<>NewValue) then
//    if (Grille.Cells[IndiceChamp,Row]=RechDom(RecupTabletteAssocie(champ),OldValue,False)) and (Grille.Cells[IndiceChamp+1,Row]<>RechDom(RecupTabletteAssocie(champ),NewValue,False)) then
      begin
      Grille.CellValues[IndiceChamp+1,Row]:=NewValue;
//      RechDom(RecupTabletteAssocie(champ),NewValue,False);
      Compteur:=Compteur+1;
      Modifier:=True;
      end;
    End;
  End
Else
  Begin
  PgiBox(TexteMessage[3],Ecran.caption);
  exit;
  End;
PgiInfo(IntToStr(Compteur)+' donnée(s) modifiée(s).',Ecran.Caption);
end;

procedure TOF_PGREAFFECTTABLE.MiseEnFormeGrille;
var i : integer;
begin
GlbSql:='';
ZoneAModifier:=False;
Grille.RowCount:=2;
Grille.ColCount:=2;
Grille.FixedRows:=1;
Grille.Titres.Clear;
TitreColonne.Clear;
Grille.Titres.Add('Etablissement'); Grille.ColAligns[0]:=tacenter; //Grille.ColEditables[0]:=False;
Grille.Titres.Add('Salarié');       Grille.ColAligns[1]:=tacenter; //Grille.ColEditables[1]:=False;
TitreColonne.Add(Pref+'_ETABLISSMENT');
TitreColonne.Add(Pref+'_SALARIE');
If GetControlText('LISTETABLE')='PAIEENCOURS' then
  Begin
  Grille.ColCount:=Grille.ColCount+2;
  GlbSql:=Pref+'_DATEDEBUT,'+Pref+'_DATEFIN,';
  Grille.Titres.Add('Paie du'); Grille.ColAligns[2]:=tacenter; //Grille.ColEditables[2]:=False;
  Grille.Titres.Add('Paie au'); Grille.ColAligns[3]:=tacenter; //Grille.ColEditables[3]:=False;
  TitreColonne.Add(Pref+'_DATEDEBUT');
  TitreColonne.Add(Pref+'_DATEFIN');
  End;
MiseEnFormeColGrille(VH_PAIE.PGLibelleOrgStat1,'TRAVAILN1');
MiseEnFormeColGrille(VH_PAIE.PGLibelleOrgStat2,'TRAVAILN2');
MiseEnFormeColGrille(VH_PAIE.PGLibelleOrgStat3,'TRAVAILN3');
MiseEnFormeColGrille(VH_PAIE.PGLibelleOrgStat4,'TRAVAILN4');
MiseEnFormeColGrille(VH_PAIE.PGLibCodeStat,'CODESTAT');
MiseEnFormeColGrille(VH_PAIE.PgLibCombo1,'LIBREPCMB1');
MiseEnFormeColGrille(VH_PAIE.PgLibCombo2,'LIBREPCMB2');
MiseEnFormeColGrille(VH_PAIE.PgLibCombo3,'LIBREPCMB3');
MiseEnFormeColGrille(VH_PAIE.PgLibCombo4,'LIBREPCMB4');

GlbSql:=Trim(GlbSql);
GlbSql:=Copy(GlbSql,1,Length(GlbSql)-1);
Grille.UpdateTitres;
{
for i:=0 to Grille.ColCount-1 do
  Grille.ColEditables[i]:=true;//(Pos('modifié',Grille.Cells[i,0])>1);
}
end;


procedure TOF_PGREAFFECTTABLE.MiseEnFormeColGrille(Champ,Suffixe : string);
begin
if Grille=nil then exit;
if (Champ <>'') AND (GetControlText('CK'+Suffixe)='X') then
  begin
  ZoneAModifier:=True;
  Grille.ColCount:=Grille.ColCount+2;
  Grille.Titres.Add(Champ+' affecté');
  Grille.Titres.Add(Champ+' modifié');
  TitreColonne.Add(Pref+'_'+Suffixe);
  TitreColonne.Add(Pref+'_'+Suffixe);
  Grille.ColFormats[Grille.ColCount-2]:='CB=PG'+Suffixe;
  Grille.ColFormats[Grille.ColCount-1]:='CB=PG'+Suffixe;
  Grille.CellValues[Grille.ColCount-2,0]:=Pref+'_'+Suffixe;
  Grille.CellValues[Grille.ColCount-1,0]:=Pref+'_'+Suffixe;
//  Grille.ColEditables[Grille.ColCount-2]:=False;
  GlbSql:=GlbSql+Pref+'_'+Suffixe+',';
  end;
end;

function TOF_PGREAFFECTTABLE.RecupTabletteAssocie(Champ: string): String;
begin
if Champ='' then result:=''
 else
  if champ='SO_PGLIBCODESTAT' then result:='PGCODESTAT'
  else
    if Pos('SO_PGLIBORGSTAT',champ)>0 then result:='PGTRAVAILN'+Copy(Champ,Length(Champ),1)
    else
      if Pos('SO_PGLIBCOMBO',champ)>0 then result:='PGLIBREPCMB'+Copy(Champ,Length(Champ),1);
end;

function TOF_PGREAFFECTTABLE.RecupNomTableAssocie(Champ: string): String;
begin
if Champ='' then result:=''
 else
  if champ='SO_PGLIBCODESTAT' then result:=Pref+'_CODESTAT'
  else
    if Pos('SO_PGLIBORGSTAT',champ)>0 then result:=Pref+'_TRAVAILN'+Copy(Champ,Length(Champ),1)
    else
      if Pos('SO_PGLIBCOMBO',champ)>0 then result:=Pref+'_LIBREPCMB'+Copy(Champ,Length(Champ),1);
end;


procedure TOF_PGREAFFECTTABLE.OnUpdate;
var
ListeTable : string;
NbMajOk : integer;
begin
  inherited;
ListeTable:=GetControlText('LISTETABLE');
NbMajOk:=0;
if (ListeTable='') then begin PGIBox(TexteMessage[1],Ecran.caption); exit; end;
if (ListeTable='RECOPIE') then
  Begin
  if PgiAsk(TexteMessage[9],Ecran.Caption)=MrNo then exit;
  RecopieTablesSalaries(NbMajOk);
  End
else
  Begin
  if PgiAsk(TexteMessage[10],Ecran.Caption)=MrNo then exit;
  if Modifier or ZoneAModifier then
    MiseAjourSelectif(NbMajOk)
  else
    begin PGIBox(TexteMessage[7],Ecran.caption); exit; end;
  End;
if NbMajOk=0 then PGIInfo(TexteMessage[8],Ecran.Caption)
else
  Begin
  PGIInfo(IntToStr(NbMajOk)+' enregistrement(s) sauvegardé(s).Traitement terminé.',Ecran.Caption);
  if (ListeTable<>'RECOPIE') then OnClickBtnActiveWhere(nil);
  End;
end;

procedure TOF_PGREAFFECTTABLE.MiseAjourSelectif(var NbMajOk : integer);
Var Row,IndiceChamp : integer;
ListeTable,St,StWhere : string;
begin
NbMajOk:=0;
ListeTable:=GetControlText('LISTETABLE');
InitMove(Grille.RowCount,'');
For Row:=1 to Grille.RowCount-1 do
  //Recherche par lignes des valeurs modifiées
  Begin
  St:=''; //Positionnement sur la prémière colonne de réaffectation
  if ListeTable='SALARIES' then IndiceChamp:=2
    else
    if ListeTable='PAIEENCOURS' then IndiceChamp:=4
     else Break;
  RecupClauseChampReaffecter(IndiceChamp,Row,St); // Recherche par Colonnes des valeurs modifiées
  if St<>'' then
    Begin
    try
    BeginTrans;
    if ListeTable='PAIEENCOURS' then
      Begin
      StWhere:=' WHERE '+Pref+'_SALARIE="'+Grille.Cells[1,Row]+'" '+
      'AND '+Pref+'_ETABLISSEMENT="'+Grille.Cells[0,Row]+'" '+
      'AND '+Pref+'_DATEFIN>="'+USDateTime(StrToDate(GetControlText('DATEDEB')))+'" '+
      'AND '+Pref+'_DATEFIN<="'+USDateTime(StrToDate(GetControlText('DATEFIN')))+'" ';
      MoveCur(False);
      NbMajOk:=NbMajOk+ExecuteSql('UPDATE '+LISTETABLE+' SET '+st+StWhere);
      St:=ConvertPrefixe(St,'PPU','PHB'); StWhere:=ConvertPrefixe(StWhere,'PPU','PHB');
      MoveCur(False);
      NbMajOk:=NbMajOk+ExecuteSql('UPDATE HISTOBULLETIN SET '+st+' '+StWhere);
      St:=ConvertPrefixe(St,'PHB','PHC'); StWhere:=ConvertPrefixe(StWhere,'PHB','PHC');
      MoveCur(False);
      NbMajOk:=NbMajOk+ExecuteSql('UPDATE HISTOCUMSAL SET '+st+' '+StWhere);
      St:=ConvertPrefixe(St,'PHC','PHA'); StWhere:=ConvertPrefixe(StWhere,'PHC','PHA');
      MoveCur(False);
      NbMajOk:=NbMajOk+ExecuteSql('UPDATE HISTOANALPAIE SET '+st+' '+StWhere);
      St:=ConvertPrefixe(St,'PHA','PCN'); StWhere:=ConvertPrefixe(StWhere,'PHA','PCN');
      MoveCur(False);
      St:=RendClauseAbsenceSalarie(St);
      if St<>'' then
        NbMajOk:=NbMajOk+ExecuteSql('UPDATE ABSENCESALARIE SET '+St+' '+StWhere);
      End
    else
      Begin
      St:='UPDATE '+LISTETABLE+' SET '+St+
      ' WHERE '+Pref+'_SALARIE="'+Grille.Cells[1,Row]+'" AND '+Pref+'_ETABLISSEMENT="'+Grille.Cells[0,Row]+'" ';
      NbMajOk:=NbMajOk+ExecuteSql(st);
      End;
    MoveCur(False);
    CommitTrans;
    Modifier:=False;
    Except
    Rollback; NbMajOk:=0;
    PGIBox ('Une erreur est survenue lors de la mise à jour des données', Ecran.caption);
    END;
    End;
  End;
FiniMove;
end;

procedure TOF_PGREAFFECTTABLE.RecopieTablesSalaries(var NbMajOk : integer);
Var
Pages : TPageControl;
StWhere,Champ : String;
begin
NbMajOk:=0;
Champ:='';
if (GetControlText('CKTRAVAILN1')='X') then Champ:=Champ+'PPU_TRAVAILN1=(SELECT PSA_TRAVAILN1 FROM SALARIES WHERE PSA_SALARIE=PPU_SALARIE),';
if (GetControlText('CKTRAVAILN2')='X') then Champ:=Champ+'PPU_TRAVAILN2=(SELECT PSA_TRAVAILN2 FROM SALARIES WHERE PSA_SALARIE=PPU_SALARIE),';
if (GetControlText('CKTRAVAILN3')='X') then Champ:=Champ+'PPU_TRAVAILN3=(SELECT PSA_TRAVAILN3 FROM SALARIES WHERE PSA_SALARIE=PPU_SALARIE),';
if (GetControlText('CKTRAVAILN4')='X') then Champ:=Champ+'PPU_TRAVAILN4=(SELECT PSA_TRAVAILN4 FROM SALARIES WHERE PSA_SALARIE=PPU_SALARIE),';
if (GetControlText('CKCODESTAT')='X')  then Champ:=Champ+'PPU_CODESTAT=(SELECT PSA_CODESTAT FROM SALARIES WHERE PSA_SALARIE=PPU_SALARIE),';
if (GetControlText('CKLIBREPCMB1')='X')then Champ:=Champ+'PPU_LIBREPCMB1=(SELECT PSA_LIBREPCMB1 FROM SALARIES WHERE PSA_SALARIE=PPU_SALARIE),';
if (GetControlText('CKLIBREPCMB2')='X')then Champ:=Champ+'PPU_LIBREPCMB2=(SELECT PSA_LIBREPCMB2 FROM SALARIES WHERE PSA_SALARIE=PPU_SALARIE),';
if (GetControlText('CKLIBREPCMB3')='X')then Champ:=Champ+'PPU_LIBREPCMB3=(SELECT PSA_LIBREPCMB3 FROM SALARIES WHERE PSA_SALARIE=PPU_SALARIE),';
if (GetControlText('CKLIBREPCMB4')='X')then Champ:=Champ+'PPU_LIBREPCMB4=(SELECT PSA_LIBREPCMB4 FROM SALARIES WHERE PSA_SALARIE=PPU_SALARIE),';

if Champ='' then Begin PgiBox(TexteMessage[6],Ecran.Caption); Exit; End;
InitMove(100,'');
Pages:=TPageControl(GetControl('Pages'));
SetControlProperty('DATEDEB','Name','PPU_DATEDEBUT');
SetControlProperty('DATEFIN','Name','PPU_DATEFIN');
StWhere:=RecupWhereCritere(PAges);
SetControlProperty('PPU_DATEDEBUT','Name','DATEDEB');
SetControlProperty('PPU_DATEFIN','Name','DATEFIN');
Try
  BeginTrans;
  Champ:=Copy(Champ,1,Length(Champ)-1)+' ';
  MoveCur(False);
  NbMajOk:=NbMajOk+ExecuteSql('UPDATE PAIEENCOURS SET '+Champ+StWhere);
  Champ:=ConvertPrefixe(Champ,'PPU','PHB'); StWhere:=ConvertPrefixe(StWhere,'PPU','PHB');
  MoveCur(False);
  NbMajOk:=NbMajOk+ExecuteSql('UPDATE HISTOBULLETIN SET '+Champ+' '+StWhere);
  Champ:=ConvertPrefixe(Champ,'PHB','PHC'); StWhere:=ConvertPrefixe(StWhere,'PHB','PHC');
  MoveCur(False);
  NbMajOk:=NbMajOk+ExecuteSql('UPDATE HISTOCUMSAL SET '+Champ+' '+StWhere);
  Champ:=ConvertPrefixe(Champ,'PHC','PHA'); StWhere:=ConvertPrefixe(StWhere,'PHC','PHA');
  MoveCur(False);
  NbMajOk:=NbMajOk+ExecuteSql('UPDATE HISTOANALPAIE SET '+Champ+' '+StWhere);
  Champ:=ConvertPrefixe(Champ,'PHA','PCN'); StWhere:=ConvertPrefixe(StWhere,'PHA','PCN');
  MoveCur(False);
  Champ:=RendClauseAbsenceSalarie(Champ);
  if Champ<>'' then
     NbMajOk:=NbMajOk+ExecuteSql('UPDATE ABSENCESALARIE SET '+Champ+' '+StWhere);
  MoveCur(False);
  CommitTrans;
Except
  Rollback; NbMajOk:=0;
  PGIBox ('Une erreur est survenue lors de la mise à jour des données', Ecran.caption);
End;
FiniMove;
end;


function TOF_PGREAFFECTTABLE.RendClauseAbsenceSalarie(StClause: String): String;
Var
Clause : string;
begin
result:='';
While (StClause<>'') do
  Begin
  Clause:=ReadTokenPipe(StClause,',');
  If Pos('LIBREPCMB',Clause)<1 then
    Begin
    if result<>'' then result:=result+',';
        result:=result+Clause;
      End;
    end;
end;


{ TOF_PGREAFFECTORGANISME }

procedure TOF_PGREAFFECTORGANISME.MiseEnFormeGrille;
begin
TitreColonne.Clear;
Grille.ColCount:=6;
TitreColonne.Add('PCT_RUBRIQUE'); Grille.Titres.Add('Rubrique');             Grille.CellValues[0,0]:='PCT_RUBRIQUE';
TitreColonne.Add('PCT_LIBELLE');  Grille.Titres.Add('Libellé');              Grille.CellValues[1,0]:='PCT_LIBELLE';
TitreColonne.Add('PCT_PREDEFINI');Grille.Titres.Add('Prédefini');            Grille.CellValues[2,0]:='PCT_PREDEFINI';
TitreColonne.Add('PCT_ORGANISME');Grille.Titres.Add('Organisme affecté');    Grille.CellValues[3,0]:='PCT_ORGANISME';
TitreColonne.Add('PCT_ORGANISME');Grille.Titres.Add('Organisme modifié');    Grille.CellValues[4,0]:='PCT_ORGANISME';
TitreColonne.Add('HISTO');        Grille.Titres.Add('modifié l''historique');Grille.CellValues[5,0]:='HISTO';
Grille.ColFormats[2]:='CB=PGPREDEFINI';
Grille.ColFormats[3]:='CB=PGTYPEORGANISME';
Grille.ColFormats[4]:='CB=PGTYPEORGANISME';
Grille.ColFormats[5]:='CB=PGOUINON';// IntToStr(Ord(csCheckbox)); //csCoche
//Grille.ColTypes[5]:='B';
{
Grille.ColAligns[0]:=tacenter;  Grille.ColEditables[0]:=True;
Grille.ColAligns[1]:=tacenter;  Grille.ColEditables[1]:=True;
Grille.ColAligns[2]:=tacenter;  Grille.ColEditables[2]:=True;
Grille.ColAligns[3]:=tacenter;  Grille.ColEditables[3]:=True;
Grille.ColAligns[4]:=tacenter;  Grille.ColEditables[4]:=True;
Grille.ColAligns[5]:=tacenter;  Grille.ColEditables[5]:=True;
}
ZoneAModifier:=True;
end;

procedure TOF_PGREAFFECTORGANISME.OnArgument(Arguments: String);
Var
Btn : TToolBarButton97 ;
Check : TCheckBox;
begin
  inherited;
//Gestionnnaire d'evenement
Btn := TToolBarButton97(GetControl('BTNACTIVEWHERE'));
if btn<>nil then Btn.Onclick:=OnClickBtnActiveWhere;
Btn := TToolBarButton97(GetControl('BTNMAJVALUE'));
if btn<>nil then Btn.Onclick:=OnClickBtnMajValue;
Check := TCheckBox(GetControl('CKHISTO'));
if Check<>nil then Check.OnClick:=OnClickHisto;
end;

procedure TOF_PGREAFFECTORGANISME.OnClickHisto(Sender: TObject);
Var
i : integer;
begin
SetControlEnabled('TDATEDEB',GetControlText('CKHISTO')='X');
SetControlEnabled('DATEDEB',GetControlText('CKHISTO')='X');
SetControlEnabled('TDATEFIN',GetControlText('CKHISTO')='X');
SetControlEnabled('DATEFIN',GetControlText('CKHISTO')='X');
For i:=1 to Grille.RowCount-1 do
  if GetControlText('CKHISTO')='X' then
      Grille.CellValues[5,i]:='OUI'
  else
      Grille.CellValues[5,i]:='NON';


end;

procedure TOF_PGREAFFECTORGANISME.OnClickBtnActiveWhere(Sender: TObject);
Var
StSql,Critere : string;
Q                                             : TQuery;
Pages                                         : TPageControl;
i                                             : integer;
HMTrad                                        : THSystemMenu;
CEG,STD,DOS                                   : Boolean;
begin
Grille.RowCount:=2; Grille.ColCount:=2;
MiseEnFormeGrille;
Pages:=TPageControl(GetControl('Pages'));
if Pages<>nil then Critere:=RecupWhereCritere(Pages)
else Critere:='';
if Critere<>'' then Critere:=Critere+' AND'
else           Critere:=Critere+' WHERE';
AccesPredefini('TOUS',CEG,STD,DOS);
If not CEG Then
 If STD Then Critere:=Critere+' PCT_PREDEFINI<>"CEG" AND'
  else
   If DOS Then Critere:=Critere+' PCT_PREDEFINI="DOS" AND';
StSql:='SELECT PCT_RUBRIQUE,PCT_LIBELLE,PCT_PREDEFINI,PCT_ORGANISME FROM COTISATION  '+
Critere+' ##PCT_PREDEFINI## PCT_NATURERUB="COT" ORDER BY PCT_RUBRIQUE';
Q:=OpenSql(StSql,True);
if Tob_Grille<>nil then begin Tob_Grille.free; Tob_Grille:=nil; end;
Tob_Grille:=Tob.create('La liste à maj',nil,-1);
Tob_Grille.LoadDetailDB('La liste à maj','','',Q,False);
For i:=0 to Tob_Grille.detail.count-1 do
  Begin
  Tob_Grille.Detail[i].AddChampSupValeur('PCT_ORGANISME_',Tob_Grille.Detail[i].GetValue('PCT_ORGANISME'));
  Tob_Grille.Detail[i].AddChampSupValeur('HISTO','OUI');
  End;
Ferme(Q);
Tob_Grille.PutGridDetail(Grille,False,False,'');
MiseEnFormeGrille;
Grille.Enabled:=ZoneAModifier;
Grille.UpdateTitres;
HMTrad.ResizeGridColumns(Grille);
Grille.Col:=5;
SetControlEnabled('BTNMAJVALUE',True);
SetControlEnabled('PNMODIF',True);
SetControlEnabled('BValider',True);
end;

procedure TOF_PGREAFFECTORGANISME.OnClickBtnMajValue(Sender: TObject);
Var
Row,Compteur : integer;
OldValue,NewValue : string;
begin
OldValue:= GetControlText('OLDVALUE');
NewValue:= GetControlText('NEWVALUE');
Compteur:=0;
if (OldValue<>'') and (NewValue<>'') then
  Begin
  For Row:=1 to Grille.RowCount-1 do
    Begin
    if (Grille.CellValues[3,Row]=OldValue) and (Grille.CellValues[4,Row]<>NewValue) then
      begin
      Grille.CellValues[4,Row]:=NewValue;
      Compteur:=Compteur+1;
      Modifier:=True;
      end;
    End;
  PgiInfo(IntToStr(Compteur)+' donnée(s) modifiée(s).',Ecran.Caption);
  End
Else
PgiBox(TexteMessage[3],Ecran.caption);
end;


procedure TOF_PGREAFFECTORGANISME.OnUpdate;
Var NbMajOk : integer;
Row : integer;
St,StWhere : string;
begin
  inherited;
NbMajOk:=0;
InitMove(Grille.RowCount,'');
For Row:=1 to Grille.RowCount-1 do
  //Recherche par lignes des valeurs modifiées
  Begin
  RecupClauseChampReaffecter(3,Row,St);
  if (St<>'') AND (Grille.Cellvalues[4,Row]<>'') then
    try
    BeginTrans;
    MoveCur(False);
    StWhere:=' WHERE PCT_NATURERUB="COT" AND ##PCT_PREDEFINI## PCT_RUBRIQUE="'+Grille.Cells[0,Row]+'" ';
    NbMajOk:=NbMajOk+ExecuteSql('UPDATE COTISATION SET '+st+StWhere+
    ' AND PCT_PREDEFINI="'+Grille.CellValues[2,Row]+'"');
    if UpperCase(Grille.Cells[5,Row])='OUI' then
      Begin
      St:=ConvertPrefixe(St,'PCT','PHB'); StWhere:=ConvertPrefixe(StWhere,'PCT','PHB');
      StWhere:=StWhere+' AND PHB_DATEFIN>="'+USDateTime(StrToDate(GetControlText('DATEDEB')))+'" '+
      'AND PHB_DATEFIN<="'+USDateTime(StrToDate(GetControlText('DATEFIN')))+'" ';
      MoveCur(False);
      NbMajOk:=NbMajOk+ExecuteSql('UPDATE HISTOBULLETIN SET '+st+' '+StWhere);
     // St:=ConvertPrefixe(St,'PHB','PHA'); StWhere:=ConvertPrefixe(StWhere,'PHB','PHA');
     // MoveCur(False);
     // NbMajOk:=NbMajOk+ExecuteSql('UPDATE HISTOANALPAIE SET '+st+' '+StWhere);
      End;
    MoveCur(False);
    CommitTrans;
    Modifier:=False;
    Except
    Rollback; NbMajOk:=0;
    PGIBox ('Une erreur est survenue lors de la mise à jour des données', Ecran.caption);
    END;
  End;
FiniMove;
if NbMajOk=0 then PGIInfo(TexteMessage[8],Ecran.Caption)
else
  Begin
  PGIInfo(IntToStr(NbMajOk)+' enregistrement(s) sauvegardé(s).Traitement terminé.',Ecran.Caption);
  OnClickBtnActiveWhere(nil);
  End;
end;

Initialization
registerclasses([TOF_PGREAFFECT,TOF_PGREAFFECTTABLE,TOF_PGREAFFECTORGANISME]);
end.
