{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 07/03/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGPARAMLISTEZL ()
Mots clefs ... : TOF;PGPARAMLISTEZL
*****************************************************************}
Unit UTofPGParamListeZL ;

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

{$ENDIF}
     forms,
     uTob,
     sysutils,
     HSysMenu,
     ComCtrls,
     HCtrls, 
     HEnt1,
     UTobDebug,
     HTB97,
     PgOutilsHistorique,
     HMsgBox, 
     UTOF ; 

Type
  TOF_PGPARAMLISTEZL = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    private
    TobChampDispo : Tob;
    GrilleDepart,GrilleDest : THGrid;
    Procedure AfficheGrille;
    Procedure BAjoutClick(Sender : TObject);
    Procedure BSuppClick(Sender : TObject);
    procedure MajDonnees;
  end ;

Implementation

procedure TOF_PGPARAMLISTEZL.OnUpdate ;
var i : Integer;
    TobMaj,T : Tob;
begin
  Inherited ;
   TobChampDispo.GetGridDetail(GrilleDest,-1,'','TYPE;NIVEAU;DONNEES;CODE;LIBELLE;INVISIBLE');
   ExecuteSQL('DELETE FROM PGPARAMAFFICHEZL');
   TobMaj := TOb.Create('Mise a jour',Nil,-1);
   T := Tob.Create('PGPARAMAFFICHEZL',TobMaj,-1);
   T.PutValue('PAZ_CODEPARAMZL','001');
   For i:= 0 to TobChampDispo.Detail.Count - 1 do
   begin
      If i > 29 then Exit;
      T.PutValue('PAZ_CHAMPDISPO'+IntToStr(i+1),TobChampDispo.Detail[i].GetValue('INVISIBLE')+TobChampDispo.Detail[i].GetValue('CODE'));
   end;
   T.InsertDB(Nil,False);
   TobMaj.Free;
   MajDonnees;
end ;

procedure TOF_PGPARAMLISTEZL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PGPARAMLISTEZL.OnArgument (S : String ) ;
var BtPlus,BtMoins : TToolBarButton97;
begin
  Inherited ;
  GrilleDepart := THGrid(GetControl('GDISPO'));
  GrilleDest := THGrid(GetControl('GPARAM'));
  AfficheGrille;
  BtPlus := TToolBarButton97(GetControl('BAJOUT'));
  If BtPlus <> Nil Then BtPlus.OnClick := BAjoutClick;
  BtMoins := TToolBarButton97(GetControl('BSUPP'));
  If BtMoins <> Nil Then BtMoins.OnClick := BSuppClick;
end ;

procedure TOF_PGPARAMLISTEZL.OnClose ;
begin
  Inherited ;
  If TobChampDispo <> Nil then FreeAndNil(TobChampDispo);
end ;

Procedure TOF_PGPARAMLISTEZL.AfficheGrille;
var TobChampPossible,TobTempDispo,TobTempParam,T,TTemp : Tob;
    Q : TQuery;
    i,x : Integer;
    Champ,Libelle,TypeChamp : String;
    HMTrad: THSystemMenu;
begin
GrilleDepart.RowCount := 2;
GrilleDest.RowCount := 2;
If TobChampDispo <> Nil then FreeAndNil(TobChampDispo);
TobChampDispo := Tob.Create('Champ du parametrage',Nil,-1);
Q := OpenSQL('SELECT * FROM PGPARAMAFFICHEZL',True);
TobTempDispo := Tob.Create('Champ temporaire',Nil,-1);
TobTempDispo.LoadDetailDB('Champ temporaire','','',Q,False);
Ferme(Q);
Q := OpenSQL('SELECT PPP_PGINFOSMODIF,PPP_LIBELLE,PPP_TYPENIVEAU,PPP_PGTYPEDONNE FROM PARAMSALARIE WHERE PPP_PGTYPEINFOLS="ZLS"',True);
TobTempParam := Tob.Create('Champ temporaire hors param',Nil,-1);
TobTempParam.LoadDetailDB('Champ temporaire hors param','','',Q,False);
Ferme(Q);

For i := 0 to TobTempDispo.Detail.Count - 1 do
begin
  For x := 1 to 30 do
  begin
    Champ := TobTempDispo.Detail[i].GetValue('PAZ_CHAMPDISPO'+IntToStr(x));
    If (Copy(Champ,1,3) = 'NAT') then TypeChamp := 'Elément national'
    else TypeChamp := 'Elément dynamique';
    If Champ <> '' then
    begin
      T := Tob.Create('Champ param',TobChampDispo,-1);
      T.AddChampSupValeur('TYPE',TypeChamp);
      T.AddChampSupValeur('NIVEAU','');
      T.AddChampSupValeur('DONNEES','');
      If (Copy(Champ,1,3) = 'NAT') then
      begin
        T.PutValue('NIVEAU','');
        T.PutValue('DONNEES','Valeur');
      end
      else
      begin
        TTemp :=  TobTempParam.FindFirst(['PPP_PGINFOSMODIF'],[Copy(Champ,4,Length(Champ))],False);
        If TTemp <> Nil then
        begin
          T.PutValue('NIVEAU',RechDom('PGNIVEAUAVDOS',TTemp.GetValue('PPP_TYPENIVEAU'),False));
          T.PutValue('DONNEES',RechDom('PGTYPEDONNEE',TTemp.GetValue('PPP_PGTYPEDONNE'),False));
        end;
        FreeAndNil(TTemp);
      end;
      T.AddChampSupValeur('CODE',Copy(Champ,4,Length(Champ)));
      If Copy(Champ,1,3) = 'NAT' then T.AddChampSupValeur('LIBELLE',RechDom('PGELEMENTNAT',Copy(Champ,4,Length(Champ)),False))
      else T.AddChampSupValeur('LIBELLE',RechDom('PGZONEHISTOSAL',Copy(Champ,4,Length(Champ)),False));
      T.AddChampSupValeur('INVISIBLE',Copy(Champ,1,3));
    end;
  end;
end;
TobTempDispo.Free;
TobChampPossible := Tob.Create('Champ hors parametrage',Nil,-1);
For i := 0 to TobTempParam.Detail.Count - 1 do
begin
  Champ := TobTempParam.Detail[i].GetValue('PPP_PGINFOSMODIF');
  Libelle := TobTempParam.Detail[i].GetValue('PPP_LIBELLE');
  T :=  TobChampDispo.FindFirst(['CODE','INVISIBLE'],[Champ,'ZLS'],False);
  If T = nil then
  begin
    T := Tob.Create('Champ param',TobChampPossible,-1);
    T.AddChampSupValeur('TYPE','Elément dynamique');
    T.AddChampSupValeur('NIVEAU',RechDom('PGNIVEAUAVDOS',TobTempParam.Detail[i].GetValue('PPP_TYPENIVEAU'),False));
    T.AddChampSupValeur('DONNEES',RechDom('PGTYPEDONNEE',TobTempParam.Detail[i].GetValue('PPP_PGTYPEDONNE'),False));
    T.AddChampSupValeur('CODE',Champ);
    T.AddChampSupValeur('LIBELLE',Libelle);
    T.AddChampSupValeur('INVISIBLE','ZLS');
  end;
end;
TobTempParam.Free;
//Element nationaux
{Q := OpenSQL('Select DISTINCT PEL_CODEELT, PEL_LIBELLE from ELTNATIONAUX Where (##PEL_PREDEFINI##) order by PEL_CODEELT',True);
TobTempParam := Tob.Create('Champ temporaire hors param',Nil,-1);
TobTempParam.LoadDetailDB('Champ temporaire hors param','','',Q,False);
Ferme(Q);
For i := 0 to TobTempParam.Detail.Count - 1 do
begin
  Champ := TobTempParam.Detail[i].GetValue('PEL_CODEELT');
  Libelle := TobTempParam.Detail[i].GetValue('PEL_LIBELLE');
  T :=  TobChampDispo.FindFirst(['CODE','INVISIBLE'],[Champ,'NAT'],False);
  If T = nil then
  begin
    T := Tob.Create('Champ param',TobChampPossible,-1);
    T.AddChampSupValeur('TYPE','Elément national');
    T.AddChampSupValeur('NIVEAU','');
    T.AddChampSupValeur('DONNEES','Valeur');
    T.AddChampSupValeur('CODE',Champ);
    T.AddChampSupValeur('LIBELLE',Libelle);
    T.AddChampSupValeur('INVISIBLE','NAT');
  end;
end;
TobTempParam.Free;          }

Q := OpenSQL('SELECT PED_CODEELT,PED_TYPENIVEAU from ELTNATIONDOS GROUP BY PED_CODEELT,PED_TYPENIVEAU order by PED_CODEELT',True);
TobTempParam := Tob.Create('Champ temporaire hors param',Nil,-1);
TobTempParam.LoadDetailDB('Champ temporaire hors param','','',Q,False);
Ferme(Q);
For i := 0 to TobTempParam.Detail.Count - 1 do
begin
  Champ := TobTempParam.Detail[i].GetValue('PED_CODEELT');
  Libelle := RechDom('PGELEMENTNAT',TobTempParam.Detail[i].GetValue('PED_CODEELT'),False);
  T :=  TobChampDispo.FindFirst(['CODE','INVISIBLE'],[Champ,'NAT'],False);
  If T = nil then
  begin
    T := Tob.Create('Champ param',TobChampPossible,-1);
    T.AddChampSupValeur('TYPE','Elément national');
    T.AddChampSupValeur('NIVEAU',RechDom('PGNIVEAUAVDOS',TobTempParam.Detail[i].GetValue('PED_TYPENIVEAU'),False));
    T.AddChampSupValeur('DONNEES','Valeur');
    T.AddChampSupValeur('CODE',Champ);
    T.AddChampSupValeur('LIBELLE',Libelle);
    T.AddChampSupValeur('INVISIBLE','NAT');
  end;
end;
TobTempParam.Free;


TobChampPossible.PutGridDetail(GrilleDepart,False,False,'TYPE;NIVEAU;DONNEES;CODE;LIBELLE;INVISIBLE',False);
TobChampDispo.PutGridDetail(GrilleDest,False,False,'TYPE;NIVEAU;DONNEES;CODE;LIBELLE;INVISIBLE',False);
TobChampPossible.Free;
GrilleDepart.ColAligns[0] := TaLeftJustify;
GrilleDepart.ColAligns[1] := TaLeftJustify;
GrilleDepart.ColAligns[2] := TaLeftJustify;
GrilleDepart.ColAligns[3] := TaLeftJustify;
GrilleDepart.ColAligns[4] := TaLeftJustify;
GrilleDest.ColAligns[0] := TaLeftJustify;
GrilleDest.ColAligns[1] := TaLeftJustify;
GrilleDest.ColAligns[2] := TaLeftJustify;
GrilleDest.ColAligns[3] := TaLeftJustify;
GrilleDest.ColAligns[4] := TaLeftJustify;

GrilleDepart.ColWidths[0] := 70;
GrilleDepart.ColWidths[1] := 70;
GrilleDepart.ColWidths[2] := 50;
GrilleDepart.ColWidths[3] := 40;
GrilleDepart.ColWidths[4] := 120;
GrilleDepart.ColWidths[5] := -1;
HMTrad.ResizeGridColumns(GrilleDepart);
GrilleDest.ColWidths[0] := 70;
GrilleDest.ColWidths[1] := 70;
GrilleDest.ColWidths[2] := 50;
GrilleDest.ColWidths[3] := 50;
GrilleDest.ColWidths[4] := 120;
GrilleDest.ColWidths[5] := -1;
HMTrad.ResizeGridColumns(GrilleDest);
end;

Procedure TOF_PGPARAMLISTEZL.BAjoutClick(Sender : TObject);
var Row,NbRow : Integer;
    Champ : String;
begin
  Row := GrilleDepart.Row;
  Champ := GrilleDepart.CellValues [3,Row];
  NbRow := GrilleDest.RowCount;
  If (NbRow = 2) and (GrilleDest.CellValues [3,NbRow-1] = '') then NbRow := 1
  else GrilleDest.Rowcount := NbRow + 1;
  GrilleDest.CellValues [0,NbRow] := GrilleDepart.CellValues [0,Row];
  GrilleDest.CellValues [1,NbRow] := GrilleDepart.CellValues [1,Row];
  GrilleDest.CellValues [2,NbRow] := GrilleDepart.CellValues [2,Row];
  GrilleDest.CellValues [3,NbRow] := Champ;
  GrilleDest.CellValues [4,NbRow] := GrilleDepart.CellValues [4,Row];
  GrilleDest.CellValues [5,NbRow] := GrilleDepart.CellValues [5,Row];;
  GrilleDepart.DeleteRow(Row);
end;

Procedure TOF_PGPARAMLISTEZL.BSuppClick(Sender : TObject);
var Row,NbRow : Integer;
    Champ : String;
begin
  Row := GrilleDest.Row;
  Champ := GrilleDest.CellValues [3,Row];
  NbRow := GrilleDepart.RowCount;
  If (NbRow = 2) and (GrilleDepart.CellValues [3,NbRow-1] = '') then NbRow := 1
  else GrilleDepart.Rowcount := NbRow + 1;
  GrilleDepart.CellValues [0,NbRow] := GrilleDest.CellValues [0,Row];
  GrilleDepart.CellValues [1,NbRow] := GrilleDest.CellValues [1,Row];
  GrilleDepart.CellValues [2,NbRow] := GrilleDest.CellValues [2,Row];
  GrilleDepart.CellValues [3,NbRow] := Champ;
  GrilleDepart.CellValues [4,NbRow] := GrilleDest.CellValues [4,Row];
  GrilleDepart.CellValues [5,NbRow] := GrilleDest.CellValues [5,Row];;
  GrilleDest.DeleteRow(Row);
end;

procedure TOF_PGPARAMLISTEZL.MajDonnees;
Var Q : TQuery;
    i,x : Integer;
    TobSal,T,TobParam : Tob;
    Champ,TypeChamp,Valeur : String;
begin
   ExecuteSQL('DELETE FROM PGTEMPZONELIBRE');
   Q := OpenSQL('SELECT PSA_SALARIE FROM SALARIES',True);
   TobSal := Tob.Create('Les Salaries',Nil,-1);
   TobSal.LoadDetailDB('Les Salaries','','',Q,False);
   Ferme(Q);
   For i := 0 to TobSal.Detail.Count - 1 do
   begin
      T := Tob.Create('PGTEMPZONELIBRE',Nil,-1);
      T.AddChampSupValeur('PTZ_SALARIE',TobSal.Detail[i].GetValue('PSA_SALARIE'),False);
      T.InsertDB(Nil,False);
   end;
   TobSal.Free;
   Q := OpenSQL('SELECT * FROM PGPARAMAFFICHEZL',True);
   TobParam := Tob.Create('MaTob',Nil,-1);
   TobParam.LoadDetailDB('Table','','',Q,False);
   Ferme(Q);
   For i := 0 to TobParam.Detail.Count - 1 do
   begin
     For x := 1 to 30 do
     begin
     Valeur := TobParam.Detail[i].GetValue('PAZ_CHAMPDISPO'+IntToStr(x));
     Champ := Copy(Valeur,4,Length(Valeur));
     TypeChamp := Copy(Valeur,1,3);
      If Champ <> '' then
      begin
        If TypeChamp = 'ZLS' then
            ExecuteSQL('UPDATE PGTEMPZONELIBRE SET PTZ_PGVALZL'+IntToStr(x)+'=(SELECT ##TOP 1## PHD_NEWVALEUR FROM PGHISTODETAIL WHERE PHD_SALARIE=PTZ_SALARIE'+
            ' AND PHD_PGINFOSMODIF="'+Champ+'" AND PHD_DATEAPPLIC<="'+UsDateTime(V_PGI.DateEntree)+'" ORDER BY PHD_DATEAPPLIC DESC)')
        else
            //Element nat
            ExecuteSQL('UPDATE PGTEMPZONELIBRE SET PTZ_PGVALZL'+IntToStr(x)+'=(SELECT ##TOP 1## PED_MONTANTEURO FROM ELTNATIONDOS'+
            ' WHERE PED_CODEELT="'+Champ+'" AND PED_VALEURNIVEAU=PTZ_SALARIE AND PED_TYPENIVEAU = "SAL" AND PED_DATEVALIDITE<="'+UsDateTime(V_PGI.DateEntree)+'"'+
            ' ORDER BY PED_DATEVALIDITE DESC)');
      end;
     end;
   end;
   TobParam.Free;
   

end;

Initialization
  registerclasses ( [ TOF_PGPARAMLISTEZL ] ) ;
end.


