{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 15/12/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGHISTODETAILCONSULT ()
Mots clefs ... : TOF;PGHISTODETAILCONSULT
*****************************************************************}
Unit UtofPGHistoDetailConsult ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     Fe_Main,
{$else}
     eMul,
     MainEAGL,
{$ENDIF}
     forms,
     HTB97,
     Vierge,
     uTob,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HSysMenu,
     PGOutilsHistorique,
     UTOF ;

Type
  TOF_PGHISTODETAILCONSULT = Class (TOF)

    procedure OnArgument (S : String ) ; override ;
    private
    ChampModif,Salarie : String;
    procedure BinsertClick(Sender : TObject);
  end ;


Implementation

procedure TOF_PGHISTODETAILCONSULT.OnArgument (S : String ) ;
var TypeChamp,Tablette : String;
    Q : TQuery;
    TobHisto,TobGrille,TG : Tob;
    Grille : THGrid;
    i : Integer;
    HMTRAD : THSystemMenu;
    DD,DF : TDateTime;
    Val,AncVal : String;
    Pct,mt,AncMt,NewMt : Double;
    Bt : TToolBarButton97;
begin
  Inherited ;
//  PGSalariesHisto;
  Salarie := ReadTokenPipe(S,';');
  ChampModif := ReadTokenPipe(S,';');
  TFVierge(Ecran).Caption := 'Consultation pour '+RechDom('PGCHAMPDISPO',ChampModif,False)+' du salarié '+RechDom('PGSALARIE',Salarie,False); 
  Q := OpenSQL('SELECT * FROM PGHISTODETAIL WHERE PHD_SALARIE="'+Salarie+'"'+
  ' AND PHD_PGINFOSMODIF="'+ChampModif+'" ORDER BY PHD_DATEAPPLIC DESC',True);
  TobHisto := Tob.Create('Historique',Nil,-1);
  TobHisto.LoadDetailDB('Historique','','',Q,False);
  Ferme(Q);
  Q := OpenSQL('SELECT PAI_LETYPE,PAI_LIENASSOC FROM PAIEPARIM WHERE PAI_IDENT="'+ChampModif+'"',True);
  If Not Q.Eof then
  begin
    TypeChamp := Q.FindField('PAI_LETYPE').AsString;
    Tablette := Q.FindField('PAI_LIENASSOC').AsString;
  end;
  Ferme(Q);

  Grille := THGrid(GetControl('GHISTO'));
  If (TypeChamp <> 'I') and (TypeChamp <> 'F') then
  begin
       Grille.ColWidths[3] := -1;
       Grille.ColWidths[4] := -1;
  end;
  If TypeChamp = 'T' then
  begin
       Grille.ColFormats[2] := 'CB='+Tablette;
  end;
  TobGrille := Tob.Create('RemplirGrille',Nil,-1);
  For i := 0 to TobHisto.Detail.Count - 1 do
  begin
       DD := TobHisto.Detail[i].GetValue('PHD_DATEAPPLIC');
//       If i < TobHisto.Detail.Count - 1 then DF := TobHisto.Detail[i+1].GetValue('PHD_DATEAPPLIC')
//       else DF := IDate1900;
       DF := TobHisto.Detail[i].GetValue('PHD_DATEFINVALID');
       Val := TobHisto.Detail[i].GetValue('PHD_NEWVALEUR');
       If i > 0 then AncVal := TobHisto.Detail[i - 1].GetValue('PHD_NEWVALEUR')
       Else AncVal := '';
       TG := Tob.Create('FilleGrille',TobGrille,-1);
       TG.AddChampSupValeur('DATEDEB',DD);
       TG.AddChampSupValeur('DATEFIN',DF);
       TG.AddChampSupValeur('VALEUR',Val);
       If (TypeChamp = 'I') or (TypeChamp = 'F') then
       begin
            If IsNumeric(Val) then  NewMt := StrToFloat(Val)
            Else NewMt := 0;
            If IsNumeric(AncVal) then AncMt := StrToFloat(AncVal)
            Else AncMt := 0;
            Mt := 0;
            Pct := 0;
            If AncMt <> 0 then Pct := (NewMt - AncMt)/ AncMt;
            Pct := Pct * 100;
            Mt := NewMt - AncMt;
            TG.AddChampSupValeur('PCTAUG',Pct);
            TG.AddChampSupValeur('MTAUG',Mt);
       end
       else
       begin
            TG.AddChampSupValeur('PCTAUG',0);
            TG.AddChampSupValeur('MTAUG',0);
       end;
  end;
  TobHisto.Free;
  TobGrille.PutGridDetail(Grille,False,False,'DATEDEB;DATEFIN;VALEUR;PCTAUG;MTAUG');
  Grille.RowCount := TobGrille.Detail.Count + 1;
  TobGrille.Free;
  HMTrad:=THSystemMenu(GetControl('HMTrad'));
  HMTrad.ResizeGridColumns(Grille) ;
  Bt := TToolBarButton97(GetControl('BNEW'));
  If Bt <> Nil then Bt.OnClick := BinsertClick;
end ;

procedure TOF_PGHISTODETAILCONSULT.BinsertClick(Sender : TObject);
begin
     AglLanceFiche('PAY','PGHISTODETAIL','','',Salarie+';'+ChampModif+';ACTION=CREATION');
end;

Initialization
  registerclasses ( [ TOF_PGHISTODETAILCONSULT ] ) ;
end.

