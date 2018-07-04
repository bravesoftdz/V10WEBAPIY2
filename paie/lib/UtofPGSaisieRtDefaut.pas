{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 23/12/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGSAISIERTFDEFAUT ()
Mots clefs ... : TOF;PGSAISIERTFDEFAUT
*****************************************************************}
{
PT1 : 21/12/2006 FC V_70 FQ 13581 Orthographe
PT2 : 12/04/2007 MF Modifs Mise en base des fichiers Ducs EDI
                    On ne traite pas les fichiers Ducs EDI
PT3 : 28/06/2007 JL Affichage nature honoraire
}
Unit UtofPGSaisieRtDefaut ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
{$else}
     eMul, 
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HSysMenu,
     htb97,
     Vierge,
     UTOF ;

Type
  TOF_PGSAISIERTFDEFAUT = Class (TOF)
     procedure OnArgument (S : String ) ; override ;
     private
     Grille : THGrid;
     NatureFic : String;
     procedure RemplirTob(Sender : TObject);
     procedure ClickGrille(Sender : TObject);
     procedure ValiderSaisie;
     procedure VerifModifs;
     procedure ClickBVal(Sender : TOBject);
  end ;

Implementation


procedure TOF_PGSAISIERTFDEFAUT.OnArgument (S : String ) ;
var Combo : THValComboBox;
    BVal,BDefaire : TToolBarButton97;
    Crit2 : String;
begin
  Inherited ;
  Crit2 := S;
  Grille :=  THGrid(getControl('GRILLEDOC'));
  Grille.ColDrawingModes[3]:= 'IMAGE';
  Grille.ColEditables[3] := False;
  Combo := THValComboBox(GetControl('TYPEFIC'));
  If Combo <> Nil then Combo.OnChange := RemplirTob;
// d PT2
  // On ne traite pas les fichiers Ducs EDI : modif propriété Plus de TYPEFIC
  if Combo <> Nil then
    Combo.Plus := ' AND CO_CODE <>  "DUA" AND '+
                  'CO_CODE <>  "DUI" AND '+
                  'CO_CODE <>  "DUU" ';
// f PT2

  Grille.OnClick := ClickGrille;
  If Crit2 = '' then SetControlText('TYPEFIC','CER')
  else
  begin
       SetControlText('TYPEFIC',Crit2);
       SetControlEnabled('TYPEFIC',False);
  end;
  BVal := TToolBarButton97(GetControl('BValider'));
  If BVal <> Nil then BVal.OnClick := ClickBVal;
  BDefaire := TToolBarButton97(GetControl('BDefaire'));
  If BDefaire <> Nil then BDefaire.OnClick := RemplirTob;
  NatureFic := 'CER';
  TFVierge(Ecran).Retour := '';
end ;

procedure TOF_PGSAISIERTFDEFAUT.RemplirTob(Sender : Tobject);
var TFic,TobGrille,T : Tob;
    Q : TQuery;
    NbDocs,i : Integer;
    HMTRAD : THSystemMenu;
    SqlPcl : String;
begin
     VerifModifs;
     NatureFic := GetControltext('TYPEFIC');
     If V_PGI.ModePcl = '1' then SQLPcl := ' AND NOT (YFS_PREDEFINI="DOS" AND YFS_NODOSSIER<>"'+V_PGI.NoDossier+'")'
     else SQLPcl := '';
     Q := OpenSQL('SELECT YFS_PREDEFINI,YFS_NOM,YFS_BCRIT1,YFS_LIBELLE FROM YFILESTD '+
     ' WHERE YFS_CODEPRODUIT="PAIE" AND (YFS_CRIT1="MAW" OR YFS_CRIT1="XLS") AND YFS_CRIT2="'+GetControlText('TYPEFIC')+'"'+SQLPcl,True); //PT3
     TFic := Tob.Create('LesFichiers',Nil,-1);
     TFic.LoadDetailDB('LesFichiers','','',Q,False);
     Ferme(Q);
     TobGrille := Tob.Create('remplirGrille',Nil,-1);
     For i := 0 to TFic.Detail.Count - 1 do
     begin
          T := Tob.Create('RemplirGrille',TobGrille,-1);
          T.AddChampSupValeur('NOM',TFic.detail[i].GetValue('YFS_NOM'));
          T.AddChampSupValeur('LIBELLE',TFic.detail[i].GetValue('YFS_LIBELLE'));
          T.AddChampSupValeur('PREDEFINI',TFic.detail[i].GetValue('YFS_PREDEFINI'));
          if TFic.detail[i].GetValue('YFS_BCRIT1') = 'X' then T.AddChampSupValeur('DEFAUT','#ICO#47')
          else T.AddChampSupValeur('DEFAUT','#ICO#46');
     end;
     TFic.Free;
     NbDocs := TobGrille.Detail.Count;
     If NbDocs <= 0 then NbDocs := 1;
     Grille.RowCount := NbDocs + 1;
     Grille.FixedRows := 1;
     Grille.ColFormats[2] := 'CB=PGPREDEFINI';
     If TobGrille.Detail.Count = 0 then
     begin
          Grille.CellValues[0,1] := '';
          Grille.CellValues[1,1] := '';
          Grille.CellValues[2,1] := '';
          Grille.CellValues[3,1] := '';
     end;
     TobGrille.PutGridDetail(Grille,False,False,'NOM;LIBELLE;PREDEFINI;DEFAUT',False);
     TobGrille.Free;
     HMTrad.ResizeGridColumns(Grille) ;
end;

procedure TOF_PGSAISIERTFDEFAUT.ClickGrille(Sender : TObject);
var Ligne,i,Colonne : Integer;
    Defaut : Boolean;
    Valeur : String;
begin
     Colonne := Grille.Col;
     If Colonne = 3 then
     begin
          Ligne := Grille.Row;
          Valeur := Grille.CellValues[Colonne,Ligne];
          If (Valeur = '#ICO#89') OR (Valeur = '#ICO#47') then Defaut := True
          else Defaut := False;
          If Defaut then Grille.CellValues[Colonne,Ligne] := '#ICO#88'
          else Grille.CellValues[Colonne,Ligne] := '#ICO#89';
          If Not Defaut then
          begin
               For i := 1 to Grille.RowCount - 1 do
               begin
                    If Ligne <> i then
                    begin
                         Valeur := Grille.CellValues[Colonne,i];
                         If (Valeur = '#ICO#89') OR (Valeur = '#ICO#47') then Grille.CellValues[Colonne,i] := '#ICO#88';
                    end;
               end;
          end;
     end;
end;

procedure TOF_PGSAISIERTFDEFAUT.ClickBVal(Sender : TObject);
begin
     NatureFic := GetControltext('TYPEFIC');
     ValiderSaisie;
end;

procedure TOF_PGSAISIERTFDEFAUT.ValiderSaisie;
var   i : Integer;
      SQLPcl : String;
begin
     TFVierge(Ecran).Retour := '';
     If V_PGI.ModePcl = '1' then SQLPcl := ' AND NOT (YFS_PREDEFINI="DOS" AND YFS_NODOSSIER<>"'+V_PGI.NoDossier+'")'
     else SQLPcl := '';
     For i := 1 to Grille.RowCount - 1 do
     begin
          If (Grille.CellValues[3,i] = '#ICO#47') or (Grille.CellValues[3,i] = '#ICO#89') then
          begin
               ExecuteSQL('UPDATE YFILESTD SET YFS_BCRIT1="X" WHERE '+
               'YFS_NOM="'+Grille.CellValues[0,i]+'" AND YFS_CRIT2="'+NatureFic+'" AND YFS_PREDEFINI="'+Grille.CellValues[2,i]+'"'+SQLPcl);
               Grille.CellValues[3,i] := '#ICO#47';
               TFVierge(Ecran).Retour := Grille.CellValues[2,i]+Grille.CellValues[0,i];
          end
          else
          begin
               ExecuteSQL('UPDATE YFILESTD SET YFS_BCRIT1="-" WHERE  '+
               'YFS_NOM="'+Grille.CellValues[0,i]+'" AND YFS_CRIT2="'+NatureFic+'" AND YFS_PREDEFINI="'+Grille.CellValues[2,i]+'"'+SQLPcl);
               Grille.CellValues[3,i] := '#ICO#46';
          end;
     end;
end;

procedure TOF_PGSAISIERTFDEFAUT.VerifModifs;
Var Rep,i : Integer;
    Modif : Boolean;
begin
     modif := False;
     For i := 1 to Grille.RowCount - 1 do
     begin
          If (Grille.CellValues[3,i] = '#ICO#88') or (Grille.CellValues[3,i] = '#ICO#89') then modif := true;
     end;
     If modif then
     begin
          Rep := PGIAsk('Voulez-vous sauvgarder les modifications',Ecran.Caption); // PT1
          If Rep = MrYes then ValiderSaisie;
     end;
end;

Initialization
  registerclasses ( [ TOF_PGSAISIERTFDEFAUT ] ) ;
end.

