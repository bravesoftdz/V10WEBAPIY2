{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 12/10/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : UTOFGCTRAITEADRESSE ()
Mots clefs ... : TOF;UTOFGCTRAITEADRESSE
*****************************************************************}
Unit UTOFGCTRAITEADRESSE;

Interface

Uses Windows, StdCtrls, Controls, Classes, Graphics, ExtCtrls,  forms, sysutils,
      ComCtrls, AGLInit, ParamSoc, HCtrls, HEnt1, HMsgBox, HTB97, UTOF, M3FP,
{$IFDEF EAGLCLIENT}
     Maineagl
{$ELSE}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_Main
{$ENDIF}
;
procedure Entree_TraiteAdresse;
procedure ADRESSES_Vers_PIECEADRESSE;

Type
  TOF_GCTRAITEADRESSE = Class (TOF)
    Image1 : TImage;
    Image2 : TImage;
    Image3 : TImage;
    Image4 : TImage;
    Image5 : TImage;
    Image6 : TImage;
    Etape1 : THLabel;
    Etape2 : THLabel;
    Etape3 : THLabel;
    Etape4 : THLabel;
    Etape5 : THLabel;
    Etape6 : THLabel;
    TENTGEN : THLabel;
    TENTACH : THLabel;
    TENTVEN : THLabel;
    MsgBox : THMsgBox;
    ListeMessage: THListBox;
    BValider : TToolbarButton97;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  private
    Erreur : boolean;
    Sql : string;
    procedure BValiderClick(Sender: TObject);
    procedure TraiteLesAdressesVente;
    procedure TraiteLesAdressesAchat;
    function  RecupChampsSql(Nature : string) : string;
    procedure PrepareExecuteSql(Nature : string; NumEtape : integer; Transact : boolean);
    function  ExecuteSqlTransaction(Sql : string) : boolean;
  end ;

Const
	TexteMessage: array[1..10] of string 	= (
          {1}     'La mise à jour des tiers livrés s''est mal passée !',
          {2}     'La mise à jour des tiers facturés s''est mal passée !',
          {3}     'La création des adresses de facturation s''est mal passée !',
          {4}     'La création des adresses de livraison s''est mal passée !',
          {5}     'La mise à jour des adresses de facturation s''est mal passée !',
          {6}     'La mise à jour des adresses de facturation s''est mal passée !',
          {7}     'La mise à jour des adresses de livraison s''est mal passée !',
          {8}     'Le traitement s''est terminé correctement !',
          {9}     'Le traitement s''est mal terminé !',
         {10}     'Le traitement a déjà été effectué !'
                  );
Implementation
uses CbpMCD
  	,CbpEnumerator
    ;

function LanceTraitementSQL(Lequel : integer) : boolean;
var Sql : string;
begin
  Result := False;
  case Lequel of
   1 : Sql := 'update piece set gp_tierslivre=gp_tiers Where gp_tierslivre=""';
   2 : Sql := 'update piece set gp_tiersfacture=gp_tiers Where gp_tiersfacture=""';
   3 : Sql := 'insert INTO pieceadresse select gp_naturepieceg as GPA_NATUREPIECEG, ' +
           'gp_souche as GPA_SOUCHE, gp_numero as GPA_NUMERO, gp_indiceg as GPA_INDICEG, ' +
           '0 as GPA_NUMLIGNE, "002" as GPA_TYPEPIECEADR, adr_juridique as GPA_JURIDIQUE, ' +
           'adr_libelle as GPA_LIBELLE, adr_libelle2 as GPA_LIBELLE2, adr_adresse1 as GPA_ADRESSE1, ' +
           'adr_adresse2 as GPA_ADRESSE2, adr_adresse3 as GPA_ADRESSE3, ' +
           'adr_codepostal as GPA_CODEPOSTAL, adr_ville as GPA_VILLE, adr_pays as GPA_PAYS ' +
           'from piece, adresses Where adr_numeroadresse=gp_numadressefact ' +
           'and gp_numadressefact<>gp_numadresselivr order by gp_numero';
   4 : Sql := 'insert INTO pieceadresse select gp_naturepieceg as GPA_NATUREPIECEG, ' +
           'gp_souche as GPA_SOUCHE, gp_numero as GPA_NUMERO, gp_indiceg as GPA_INDICEG, ' +
           '0 as GPA_NUMLIGNE, "001" as GPA_TYPEPIECEADR, adr_juridique as GPA_JURIDIQUE, ' +
           'adr_libelle as GPA_LIBELLE, adr_libelle2 as GPA_LIBELLE2, adr_adresse1 as GPA_ADRESSE1, ' +
           'adr_adresse2 as GPA_ADRESSE2, adr_adresse3 as GPA_ADRESSE3, ' +
           'adr_codepostal as GPA_CODEPOSTAL, adr_ville as GPA_VILLE, adr_pays as GPA_PAYS ' +
           'from piece, adresses Where adr_numeroadresse=gp_numadresselivr ' +
           'order by gp_numero';
   5 : Sql := 'update piece set gp_numadresseFact=2 Where gp_numadressefact<>gp_numadresselivr';
   6 : Sql := 'update piece set gp_numadresseFact=1 Where gp_numadressefact=gp_numadresselivr';
   7 : Sql := 'update piece set gp_numadresselivr=1';
   end;

  BeginTrans;
  Try
    ExecuteSql(Sql);
    CommitTrans;
    Result:=True;
  Except
    RollBack ;
    end;
end;

procedure ADRESSES_Vers_PIECEADRESSE;
var i : integer;
    TraitementOk : boolean;
begin
if Not GetParamSoc('SO_GCPIECEADRESSE') then
  begin
  TraitementOk := True;
  For i:=1 to 7 do
    begin
    if Not LanceTraitementSQL(i) then begin TraitementOk:=False; Break; end;
    end;
  if TraitementOk then SetParamSoc('SO_GCPIECEADRESSE','X');
  end;
end;

procedure Entree_TraiteAdresse;
begin
if GetParamSoc('SO_GCPIECEADRESSE') then
   begin
   PGIBox(TexteMessage[10],'Basculement adresses');
   Exit;
   end;
V_PGI.ZoomOLE := True;
AGLLanceFiche('GC','GCTRAITEADRESSE','','','');
V_PGI.ZoomOLE := False;
end;

procedure TOF_GCTRAITEADRESSE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_GCTRAITEADRESSE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_GCTRAITEADRESSE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_GCTRAITEADRESSE.OnLoad ;
var i_ind1, i_ind2 : integer;
    st1, st2 : string;
begin
  Inherited ;
  Image1 := TImage(GetControl('IMAGE1'));
  Image2 := TImage(GetControl('IMAGE2'));
  Image3 := TImage(GetControl('IMAGE3'));
  Image4 := TImage(GetControl('IMAGE4'));
  Image5 := TImage(GetControl('IMAGE5'));
  Image6 := TImage(GetControl('IMAGE6'));
  Etape1 := THLabel(GetControl('TETAPE1'));
  Etape2 := THLabel(GetControl('TETAPE2'));
  Etape3 := THLabel(GetControl('TETAPE3'));
  Etape4 := THLabel(GetControl('TETAPE4'));
  Etape5 := THLabel(GetControl('TETAPE5'));
  Etape6 := THLabel(GetControl('TETAPE6'));
  TENTGEN := THLabel(GetControl('TENTGEN'));
  TENTACH := THLabel(GetControl('TENTACH'));
  TENTVEN := THLabel(GetControl('TENTVEN'));

  Image1.Visible := False;
  Image2.Visible := False;
  Image3.Visible := False;
  Image4.Visible := False;
  Image5.Visible := False;
  Image6.Visible := False;
  Etape1.Font.Style := Etape1.Font.Style - [fsBold, fsItalic];
  Etape2.Font.Style := Etape2.Font.Style - [fsBold, fsItalic];
  Etape3.Font.Style := Etape3.Font.Style - [fsBold, fsItalic];
  Etape4.Font.Style := Etape4.Font.Style - [fsBold, fsItalic];
  Etape5.Font.Style := Etape5.Font.Style - [fsBold, fsItalic];
  Etape6.Font.Style := Etape6.Font.Style - [fsBold, fsItalic];
  TENTGEN.Visible := True;
  TENTACH.Visible := False;
  TENTVEN.Visible := False;

  MsgBox := THMsgBox.Create(Ecran);
  ListeMessage := ThListBox(GetControl('LISTEMESSAGE'));
  for i_ind1 := 0 to ListeMessage.Items.Count - 1 do
      begin
      st1 := ListeMessage.Items.Strings[i_ind1];
      if IsNumeric(st1[1]) then
          begin
          for i_ind2 := i_ind1 + 1 to ListeMessage.Items.Count - 1 do
              begin
              st2 := ListeMessage.Items.Strings[i_ind2];
              if IsNumeric(st2[1]) then break;
              st1 := st1 + st2;
              end;
          ListeMessage.Items.Strings[i_ind1] := st1;
          end;
      end;
  for i_ind1 := ListeMessage.Items.Count - 1 downto 0 do
      begin
      st1 := ListeMessage.Items.Strings[i_ind1];
      if not IsNumeric(st1[1]) then ListeMessage.Items.Delete(i_ind1);
      end;
  MsgBox.Mess := ListeMessage.Items;
  BValider := TToolbarButton97(GetControl('BValider'));
  BValider.OnClick := BValiderClick;
end ;

procedure TOF_GCTRAITEADRESSE.OnArgument (S : String ) ;
begin
  Inherited ;
end ;

procedure TOF_GCTRAITEADRESSE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_GCTRAITEADRESSE.BValiderClick(Sender: TObject);
var i_ind1 : integer;
begin
  inherited;
for i_ind1 := 0 to MsgBox.Mess.Count - 1 do
    if MsgBox.Execute (i_ind1,Ecran.Caption,'') <> mrYes then Exit;
Erreur := False;
TraiteLesAdressesVente;
if not Erreur then TraiteLesAdressesAchat;
if not Erreur then
    begin
    SetParamSoc('SO_GCPIECEADRESSE','X');
    LastError:=8 ; LastErrorMsg:=TexteMessage[LastError];
    AfficheError:=False;
    PGIInfo(TexteMessage[LastError],TForm(Ecran).Caption);
    end
    else
    begin
    LastError:=9 ; LastErrorMsg:=TexteMessage[LastError];
    AfficheError:=False;
    PGIBox(TexteMessage[LastError],TForm(Ecran).Caption);
    end;
end;

procedure TOF_GCTRAITEADRESSE.TraiteLesAdressesVente;
begin
TENTGEN.Visible := False;
TENTVEN.Visible := True;
Etape1.Font.Style := Etape1.Font.Style + [fsBold, fsItalic];
PrepareExecuteSql('VEN', 1, False);
if Erreur then Exit;
Etape1.Font.Style := Etape1.Font.Style - [fsBold, fsItalic];
Image1.Visible := True;
Etape2.Font.Style := Etape2.Font.Style + [fsBold, fsItalic];
PrepareExecuteSql('VEN', 2, False);
if Erreur then Exit;
Etape2.Font.Style := Etape2.Font.Style - [fsBold, fsItalic];
Image2.Visible := True;
Etape3.Font.Style := Etape3.Font.Style + [fsBold, fsItalic];
PrepareExecuteSql('VEN', 3, False);
if Erreur then Exit;
Etape3.Font.Style := Etape3.Font.Style - [fsBold, fsItalic];
Image3.Visible := True;
Etape4.Font.Style := Etape4.Font.Style + [fsBold, fsItalic];
PrepareExecuteSql('VEN', 4, False);
if Erreur then Exit;
Etape4.Font.Style := Etape4.Font.Style - [fsBold, fsItalic];
Image4.Visible := True;
Etape5.Font.Style := Etape5.Font.Style + [fsBold, fsItalic];
PrepareExecuteSql('VEN', 5, True);
if Erreur then Exit;
PrepareExecuteSql('VEN', 6, True);
if Erreur then Exit;
Etape5.Font.Style := Etape5.Font.Style - [fsBold, fsItalic];
Image5.Visible := True;
Etape6.Font.Style := Etape6.Font.Style + [fsBold, fsItalic];
PrepareExecuteSql('VEN', 7, True);
if Erreur then Exit;
Etape6.Font.Style := Etape6.Font.Style - [fsBold, fsItalic];
Image6.Visible := True;
Image6.Invalidate;
Etape6.Invalidate;
//executeSQL('select Count(*) from piece');
end;

procedure TOF_GCTRAITEADRESSE.TraiteLesAdressesAchat;
begin
Sleep(8000);
Image1.Visible := False;
Image2.Visible := False;
Image3.Visible := False;
Image4.Visible := False;
Image5.Visible := False;
Image6.Visible := False;
Etape1.Font.Style := Etape1.Font.Style - [fsBold, fsItalic];
Etape2.Font.Style := Etape2.Font.Style - [fsBold, fsItalic];
Etape3.Font.Style := Etape3.Font.Style - [fsBold, fsItalic];
Etape4.Font.Style := Etape4.Font.Style - [fsBold, fsItalic];
Etape5.Font.Style := Etape5.Font.Style - [fsBold, fsItalic];
Etape6.Font.Style := Etape6.Font.Style - [fsBold, fsItalic];

TENTVEN.Visible := False;
TENTACH.Visible := True;
Etape1.Font.Style := Etape1.Font.Style + [fsBold, fsItalic];
PrepareExecuteSql('ACH', 1, False);
if Erreur then Exit;
Etape1.Font.Style := Etape1.Font.Style - [fsBold, fsItalic];
Image1.Visible := True;
Etape2.Font.Style := Etape2.Font.Style + [fsBold, fsItalic];
PrepareExecuteSql('ACH', 2, False);
if Erreur then Exit;
Etape2.Font.Style := Etape2.Font.Style - [fsBold, fsItalic];
Image2.Visible := True;
Etape3.Font.Style := Etape3.Font.Style + [fsBold, fsItalic];
PrepareExecuteSql('ACH', 3, False);
if Erreur then Exit;
Etape3.Font.Style := Etape3.Font.Style - [fsBold, fsItalic];
Image3.Visible := True;
Etape4.Font.Style := Etape4.Font.Style + [fsBold, fsItalic];
PrepareExecuteSql('ACH', 4, False);
if Erreur then Exit;
Etape4.Font.Style := Etape4.Font.Style - [fsBold, fsItalic];
Image4.Visible := True;
Etape5.Font.Style := Etape5.Font.Style + [fsBold, fsItalic];
PrepareExecuteSql('ACH', 5, True);
if Erreur then Exit;
PrepareExecuteSql('ACH', 6, True);
if Erreur then Exit;
Etape5.Font.Style := Etape5.Font.Style - [fsBold, fsItalic];
Image5.Visible := True;
Etape6.Font.Style := Etape6.Font.Style + [fsBold, fsItalic];
PrepareExecuteSql('ACH', 7, False);
if Erreur then Exit;
Etape6.Font.Style := Etape6.Font.Style - [fsBold, fsItalic];
Image6.Visible := True;
end;

function TOF_GCTRAITEADRESSE.RecupChampsSql(Nature : string) : string;

	function GetNbFields(Table : ItableCom) : Integer;
	var FieldList : IEnumerator ;
  begin
		Result := 0;
    if not Assigned(Table) then Exit;
    FieldList := Table.Fields as IEnumerator;
    FieldList.Reset;
    While FieldList.MoveNext do Inc(Result);
  end;

var ind1, ind2, iTemp, NumTable : integer;
    stSuffixe, stSep,StType : string;
    NumChamps : array of integer;
    NomChamps : array of string;
	Mcd : IMCDServiceCOM;
	Table     : ITableCOM ;
	FieldList : IEnumerator ;
  MaxChamps : Integer;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();

  Table := mcd.GetTable(Mcd.PrefixeToTable('GPA'));
  MaxChamps := GetNbFields(Table);
  SetLength(NumChamps, MaxChamps);
  SetLength(NomChamps, MaxChamps);
  FieldList := Table.Fields as IEnumerator;
  FieldList.Reset;
  While FieldList.MoveNext do 
  begin
    NumChamps[ind1] := (FieldList.Current as IFieldCOM).Numero;
    NomChamps[ind1] := (FieldList.Current as IFieldCOM).name;
  end;

  for ind1 := 1 to High(NumChamps) do
  begin
  //    if ind1 < High(Champs) then
    begin
      for ind2 := ind1 to High(NumChamps) do
      begin
        if NumChamps[ind1] > NumChamps[ind2] then
        begin
          iTemp := NumChamps[ind2];
          NumChamps[ind2] := NumChamps[ind1];
          NumChamps[ind1] := iTemp;
          stSep := NomChamps[ind2];
          NomChamps[ind2] := NomChamps[ind1];
          NomChamps[ind1] := stSep;
        end;
      end;
    end;
  end;

  Result := 'insert INTO pieceadresse select ';
  stSep := '';
  for ind1 := 1 to High(NomChamps) do
  begin
    stSuffixe := Copy(NomChamps[ind1], Pos('_', NomChamps[ind1]) + 1, 255);
    if (stSuffixe = 'NATUREPIECEG') or
       (stSuffixe = 'SOUCHE') or
       (stSuffixe = 'NUMERO') or
       (stSuffixe = 'INDICEG') then
        Result := Result + stSep + 'GP_' + stSuffixe + ' as ' + NomChamps[ind1]
    else if (stSuffixe = 'NUMLIGNE') then
        Result := Result + stSep + '0 as ' + NomChamps[ind1]
    else if (stSuffixe = 'TYPEPIECEADR') then
        Result := Result + stSep + '"' + Nature + '" as ' + NomChamps[ind1]
    else if (stSuffixe = 'JURIDIQUE') or
            (stSuffixe = 'LIBELLE') or
            (stSuffixe = 'LIBELLE2') or
            (stSuffixe = 'ADRESSE1') or
            (stSuffixe = 'ADRESSE2') or
            (stSuffixe = 'ADRESSE3') or
            (stSuffixe = 'CODEPOSTAL') or
            (stSuffixe = 'VILLE') or
            (stSuffixe = 'EAN') or
            (stSuffixe = 'INCOTERM') or
            (stSuffixe = 'LIEUDISPO') or
            (stSuffixe = 'MODEEXP') or
            (stSuffixe = 'NIF') or
            (stSuffixe = 'NUMEROCONTACT') or
            (stSuffixe = 'REGION') or
{CRM_20090514_CD_010;15187_DEB}
            (stSuffixe = 'TELEPHONE') or
            (stSuffixe = 'FAX') or
            (stSuffixe = 'TELEX') or
            (stSuffixe = 'LOCALTAX') or
{CRM_20090514_CD_010;15187_FIN}
            (stSuffixe = 'PAYS') then
        Result := Result + stSep + 'ADR_' + stSuffixe + ' as ' + NomChamps[ind1]
    else
    begin
        stType := MCD.Getfield(NomChamps[ind1]).Tipe;

        if stType = 'BOOLEAN' then
            Result := Result + stSep + '"-" as ' + NomChamps[ind1]
        else if (stType = 'DATE') then
            Result := Result + stSep + '"'+DateToStr(iDate1900)+'" as ' + NomChamps[ind1]
        else if (stType = 'DOUBLE') or
                (stType = 'RATE') or
                (stType = 'INTEGER') then
            Result := Result + stSep + '0 as ' + NomChamps[ind1]
        else
            Result := Result + stSep + '"" as ' + NomChamps[ind1];
    end;
    stSep := ',';
  end;
  Result := Result + ' ';
end;

procedure TOF_GCTRAITEADRESSE.PrepareExecuteSql(Nature : string; NumEtape : integer;
                                                Transact : boolean);
var
    WhereAch, WhereVen : string;

begin
WhereVen := 'where (gp_naturepieceg="DE" or gp_naturepieceg="CC" or gp_naturepieceg="CCE" ' +
{$ifdef AFFAIRE}
            'or gp_naturepieceg="FPR" or gp_naturepieceg="APR" or gp_naturepieceg="FRE" ' +   //mcd 13/02/03
{$endif}
{$ifdef BTP}
            'or gp_naturepieceg="DBT" or gp_naturepieceg="ABT" or gp_naturepieceg="FBT" ' +   //brl
            'or gp_naturepieceg="ETU" or gp_naturepieceg="PBT" or gp_naturepieceg="CBT" ' +
            'or gp_naturepieceg="DAC" or gp_naturepieceg="FRC" OR GP_NATUREPIECEG="LBT" ' +
{$endif}
            'or gp_naturepieceg="CCR" or gp_naturepieceg="PRE" or gp_naturepieceg="BLC" ' +
            'or gp_naturepieceg="LCE" or gp_naturepieceg="LCR" or gp_naturepieceg="FAC" ' +
            'or gp_naturepieceg="AVC" or gp_naturepieceg="AVS")';
WhereAch := 'where (gp_naturepieceg="ALF" or gp_naturepieceg="BLF" or gp_naturepieceg="CF" ' +
            'or gp_naturepieceg="CFR" or gp_naturepieceg="FCF" or gp_naturepieceg="FF" ' +
            'or gp_naturepieceg="FRF" or gp_naturepieceg="LFR")';
//Sleep(2000);
if (Nature = 'VEN') and (NumEtape = 1) then
    Sql := 'update piece set gp_tierslivre=gp_tiers ' + WhereVen + ' and gp_tierslivre=""'
else if (Nature = 'VEN') and (NumEtape = 2) then
    Sql := 'update piece set gp_tiersfacture=gp_tiers ' + WhereVen + ' and gp_tiersfacture=""'
else if (Nature = 'VEN') and (NumEtape = 3) then
{    Sql := 'insert INTO pieceadresse select gp_naturepieceg as GPA_NATUREPIECEG, ' +
           'gp_souche as GPA_SOUCHE, gp_numero as GPA_NUMERO, gp_indiceg as GPA_INDICEG, ' +
           '0 as GPA_NUMLIGNE, "002" as GPA_TYPEPIECEADR, adr_juridique as GPA_JURIDIQUE, ' +
           'adr_libelle as GPA_LIBELLE, adr_libelle2 as GPA_LIBELLE2, adr_adresse1 as GPA_ADRESSE1, ' +
           'adr_adresse2 as GPA_ADRESSE2, adr_adresse3 as GPA_ADRESSE3, ' +
           'adr_codepostal as GPA_CODEPOSTAL, adr_ville as GPA_VILLE, adr_pays as GPA_PAYS ' + }
    Sql := RecupChampsSql('002') +
           'from piece, adresses ' + WhereVen + ' and adr_numeroadresse=gp_numadressefact ' +
           'and gp_numadressefact<>gp_numadresselivr order by gp_numero'
else if (Nature = 'VEN') and (NumEtape = 4) then
{    Sql := 'insert INTO pieceadresse select gp_naturepieceg as GPA_NATUREPIECEG, ' +
           'gp_souche as GPA_SOUCHE, gp_numero as GPA_NUMERO, gp_indiceg as GPA_INDICEG, ' +
           '0 as GPA_NUMLIGNE, "001" as GPA_TYPEPIECEADR, adr_juridique as GPA_JURIDIQUE, ' +
           'adr_libelle as GPA_LIBELLE, adr_libelle2 as GPA_LIBELLE2, adr_adresse1 as GPA_ADRESSE1, ' +
           'adr_adresse2 as GPA_ADRESSE2, adr_adresse3 as GPA_ADRESSE3, ' +
           'adr_codepostal as GPA_CODEPOSTAL, adr_ville as GPA_VILLE, adr_pays as GPA_PAYS ' + }
    Sql := RecupChampsSql('001') +
           'from piece, adresses ' + WhereVen + ' and adr_numeroadresse=gp_numadresselivr ' +
           'order by gp_numero'
else if (Nature = 'VEN') and (NumEtape = 5) then
    Sql := 'update piece set gp_numadresseFact=2 ' + WhereVen +
           ' and gp_numadressefact<>gp_numadresselivr'
else if (Nature = 'VEN') and (NumEtape = 6) then
    Sql := 'update piece set gp_numadresseFact=1 ' + WhereVen +
           ' and gp_numadressefact=gp_numadresselivr'
else if (Nature = 'VEN') and (NumEtape = 7) then
    Sql := 'update piece set gp_numadresselivr=1 ' + WhereVen
else if (Nature = 'ACH') and (NumEtape = 1) then
    Sql := 'update piece set gp_tierslivre=gp_tiers ' + WhereAch + ' and gp_tierslivre=""'
else if (Nature = 'ACH') and (NumEtape = 2) then
    Sql := 'update piece set gp_tiersfacture=gp_tiers ' + WhereAch + ' and gp_tiersfacture=""'
else if (Nature = 'ACH') and (NumEtape = 3) then
{    Sql := 'insert INTO pieceadresse select gp_naturepieceg as GPA_NATUREPIECEG, ' +
           'gp_souche as GPA_SOUCHE, gp_numero as GPA_NUMERO, gp_indiceg as GPA_INDICEG, ' +
           '0 as GPA_NUMLIGNE, "002" as GPA_TYPEPIECEADR, adr_juridique as GPA_JURIDIQUE, ' +
           'adr_libelle as GPA_LIBELLE, adr_libelle2 as GPA_LIBELLE2, adr_adresse1 as GPA_ADRESSE1, ' +
           'adr_adresse2 as GPA_ADRESSE2, adr_adresse3 as GPA_ADRESSE3, ' +
           'adr_codepostal as GPA_CODEPOSTAL, adr_ville as GPA_VILLE, adr_pays as GPA_PAYS ' + }
     Sql := RecupChampsSql('002') +
          'from piece, adresses ' + WhereAch + ' and adr_numeroadresse=gp_numadressefact ' +
           'and gp_numadressefact<>gp_numadresselivr order by gp_numero'
else if (Nature = 'ACH') and (NumEtape = 4) then
{    Sql := 'insert INTO pieceadresse select gp_naturepieceg as GPA_NATUREPIECEG, ' +
           'gp_souche as GPA_SOUCHE, gp_numero as GPA_NUMERO, gp_indiceg as GPA_INDICEG, ' +
           '0 as GPA_NUMLIGNE, "001" as GPA_TYPEPIECEADR, adr_juridique as GPA_JURIDIQUE, ' +
           'adr_libelle as GPA_LIBELLE, adr_libelle2 as GPA_LIBELLE2, adr_adresse1 as GPA_ADRESSE1, ' +
           'adr_adresse2 as GPA_ADRESSE2, adr_adresse3 as GPA_ADRESSE3, ' +
           'adr_codepostal as GPA_CODEPOSTAL, adr_ville as GPA_VILLE, adr_pays as GPA_PAYS ' + }
    Sql := RecupChampsSql('001') +
           'from piece, adresses ' + WhereAch + ' and adr_numeroadresse=gp_numadresselivr ' +
           'order by gp_numero'
else if (Nature = 'ACH') and (NumEtape = 5) then
    Sql := 'update piece set gp_numadresseFact=2 ' + WhereAch +
           ' and gp_numadressefact<>gp_numadresselivr'
else if (Nature = 'ACH') and (NumEtape = 6) then
    Sql := 'update piece set gp_numadresseFact=1 ' + WhereAch +
           ' and gp_numadressefact=gp_numadresselivr'
else if (Nature = 'ACH') and (NumEtape = 7) then
    Sql := 'update piece set gp_numadresselivr=1 ' + WhereAch
    ;

if Transact then
    begin
    Try
        BeginTrans ;
        if ExecuteSqlTransaction(Sql) then CommitTrans
        else raise Exception.create('') ;
    Except
        Erreur := True;
        LastError:=NumEtape ; LastErrorMsg:=TexteMessage[LastError];
        AfficheError:=False;
        PGIBox(TexteMessage[LastError],TForm(Ecran).Caption);
        RollBack ;
    end;
    end
    else
    begin
    ExecuteSqlTransaction(Sql);
    end ;
end;

function TOF_GCTRAITEADRESSE.ExecuteSqlTransaction(Sql : string) : boolean;
begin
Result := True;
ExecuteSql(Sql);
end;

procedure AGLEntree_TraiteAdresse ( Parms : array of variant ; nb : integer ) ;
BEGIN
Entree_TraiteAdresse;
END ;

Initialization
  registerclasses ( [ TOF_GCTRAITEADRESSE ] ) ;
  RegisterAglProc( 'Entree_TraiteAdresse', False , 0, AGLEntree_TraiteAdresse);
end.
