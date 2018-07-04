{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 15/04/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGCOMPTEURSDIF ()
Mots clefs ... : TOF;PGCOMPTEURSDIF
*****************************************************************}
Unit UTofPGCompteursDIF ;

Interface

Uses StdCtrls, 
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     FE_Main,
{$else}
     eMul,
     MainEAGL,
{$ENDIF}
     forms, 
     sysutils,
     UTOB,
     ComCtrls,
     ParamSoc,
     HCtrls,
     HEnt1,
     HMsgBox,
     HSysMenu,
     Vierge,
     Graphics,
     Grids,
     HTB97,
     UTOF ;

Type
  TOF_PGCOMPTEURSDIF = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    private
    LeSalarie : String;
    LaDate : TDateTime;
    GrilleDIF : THGrid;
    TobEditDIF : Tob;
    procedure LancerDIF;
    procedure RemplirDroits;
    procedure RemplirFormations;
    Procedure RemplirDemande;
    Procedure CalculerTotal;
    procedure GrilleGetCellCanvas( ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
    procedure BtClick (Sender : TObject);
    procedure GrilleDblClick(Sender : TObject);
  end ;

Implementation

procedure TOF_PGCOMPTEURSDIF.LancerDIF;
var StChampGrid : String;
begin
  StChampGrid := 'SALARIE;LIBELLE;DRACQ;;FTT;FHTT;FTOTAL;;DETT;DEHTT;DETOTAL;DEETAT;;SOLDE;SOLDETHEO;TYPE;NUM;STAGE;MILLESIME;ETAB';
  TobEditDIF := Tob.Create('EditionDIF',Nil,-1);
  RemplirDroits;
  RemplirFormations;
  RemplirDemande;
  If TobEditDIF.Detail.Count = 0 then
  begin
       PGIBox('Aucun compteur pour ce salarié',Ecran.Caption);
  end;
  CalculerTotal;
  GrilleDIF.RowCount := TobEditDIF.Detail.Count + 2;
  TobEditDIF.Detail.Sort('SALARIE;ORDRE;ANNEE;DATEDEB');
  TobEditDIF.PutGridDetail(GrilleDIF,False,False,StChampGrid,False);
  TobEditDIF.Free;
end ;

procedure TOF_PGCOMPTEURSDIF.OnArgument (S : String ) ;
var HMTRAD : THSystemMenu;
    Bt : TToolBarButton97;
    StDate : String;
begin
  Inherited ;
    Bt := TToolBarButton97(GetControl('BCUMUL'));
    If Bt <> Nil then
    begin
         Bt.OnClick := BtClick;
         Bt.Down := True;
    end;
    Bt := TToolBarButton97(GetControl('BDEMANDE'));
    If Bt <> Nil then
    begin
         Bt.OnClick := BtClick;
         Bt.Down := True;
    end;
    Bt := TToolBarButton97(GetControl('BFORMATION'));
    If Bt <> Nil then
    begin
         Bt.OnClick := BtClick;
         Bt.Down := True;
    end;
    LeSalarie := ReadTokenPipe(S,';');
    StDate := ReadTokenPipe(S,';');
    If IsValidDate(StDate) then LaDate := StrToDate(StDate)
    Else LaDate := V_PGI.DateEntree;
    TFVierge(ECran).Caption := 'Suivi compteurs DIF du salarié : '+LeSalarie + ' ' + RechDom('PGSALARIE',LeSalarie,False);
    GrilleDIF := THGrid(GetControl('GDIF'));
    GrilleDIF.GetCellCanvas := GrilleGetCellCanvas;
    GrilleDIF.OnDblClick := GrilleDblClick;
    LancerDIF;
    //    0Salarié
    GrilleDIF.ColWidths[0]:= -1;
    //    1Libellé
    GrilleDIF.ColWidths[1] := 110;
    //    2Acquis
    GrilleDIF.ColFormats[2] := '#,##0.00;; ;';
    GrilleDIF.ColAligns[2] := taRightJustify;

    //   3Acquis EC
    GrilleDIF.ColWidths[3]:= 0;
    //    5Heures T
    GrilleDIF.ColFormats[4] := '#,##0.00;; ;';
    GrilleDIF.ColAligns[4] := taRightJustify;
    //    6Heures NT
    GrilleDIF.ColFormats[5] := '#,##0.00;; ;';
    GrilleDIF.ColAligns[5] := taRightJustify;
    //    7Total
    GrilleDIF.ColFormats[6] := '#,##0.00;; ;';
    GrilleDIF.ColAligns[6] := taRightJustify;
    GrilleDIF.ColWidths[7]:= 0;
    //    9Demande T
    GrilleDIF.ColFormats[8] := '#,##0.00;; ;';
    GrilleDIF.ColAligns[8] := taRightJustify;
    //    10Demande NT
    GrilleDIF.ColFormats[9] := '#,##0.00;; ;';
    GrilleDIF.ColAligns[9] := taRightJustify;
    //    11Demande total
    GrilleDIF.ColFormats[10] := '#,##0.00;; ;';
    GrilleDIF.ColAligns[10] := taRightJustify;
    //    12tat demande
    GrilleDIF.ColFormats[11] := 'CB=PGETATVALIDATION';
    GrilleDIF.ColWidths[12]:= 0;
    //    14Solde
    GrilleDIF.ColFormats[13] := '#,##0.00;; ;';
    GrilleDIF.ColAligns[13] := taRightJustify;
    //    15Solde théo
    GrilleDIF.ColFormats[14] := '#,##0.00;; ;';
    GrilleDIF.ColAligns[14] := taRightJustify;
    GrilleDIF.ColWidths[15]:= -1;
    GrilleDIF.ColWidths[16]:= -1;
    GrilleDIF.ColWidths[17]:= -1;
    GrilleDIF.ColWidths[18]:= -1;
    GrilleDIF.ColWidths[19]:= -1;
    GrilleDIF.CellValues[0,1] := '';
    GrilleDIF.CellValues[1,1] := '';
    GrilleDIF.CellValues[2,1] := '';
    GrilleDIF.CellValues[3,1] := '';
    GrilleDIF.CellValues[4,1] := 'Sur TT';
    GrilleDIF.CellValues[5,1] := 'Hors TT';
    GrilleDIF.CellValues[6,1] := 'Total';
    GrilleDIF.CellValues[7,1] := '';
    GrilleDIF.CellValues[8,1] := 'Sur TT';
    GrilleDIF.CellValues[9,1] := 'Hors TT';
    GrilleDIF.CellValues[10,1] := 'Total';
    GrilleDIF.CellValues[11,0] := '';
    GrilleDIF.CellValues[12,1] := '';
    GrilleDIF.CellValues[13,1] := '';
    GrilleDIF.CellValues[14,1] := 'Prévisionnel';
    HMTrad := THSystemMenu(GetControl('HMTRAD'));
    HMTrad.ResizeGridColumns(GrilleDIF);
end ;

procedure TOF_PGCOMPTEURSDIF.RemplirDroits;
var T,TCumul : tob;
    Cumul : String;
    Q : TQuery;
    i,AnneeI : Integer;
    aa, mm, dd : Word;
begin
     Cumul := GetParamSOc('SO_PGCUMULDIFACQUIS');
     Q := OpenSQL('SELECT YEAR(PHC_DATEDEBUT) ANNEE,SUM(PHC_MONTANT) ACQUIS '+
     'FROM HISTOCUMSAL WHERE PHC_DATEFIN < "'+UsDateTime(LaDate)+'" AND PHC_CUMULPAIE="'+Cumul+'" AND PHC_SALARIE="'+LeSalarie+'" GROUP BY YEAR(PHC_DATEDEBUT)',True);
     TCumul := Tob.Create('LesCumuls',Nil,-1);
     TCumul.LoadDetailDB('LesCumuls','','',Q,False);
     Ferme(Q);
     For i := 0 To TCumul.Detail.Count -1 do
     begin
          AnneeI := StrToInt(TCumul.Detail[i].GetValue('ANNEE'));
          T := Tob.Create('FilleEdition',TobEditDIF,-1);
          T.AddChampSupValeur('SALARIE',LeSalarie);
          T.AddChampSupValeur('TYPE','DROITS');
          T.AddChampSupValeur('ORDRE',1);
          decodedate(LaDate, aa, mm, dd);
          If aa <> AnneeI then
          begin
               T.AddChampSupValeur('LIBELLE','Cumul au '  + DateToStr(EncodeDate(AnneeI,12,31)));
               T.AddChampSupValeur('ANNEE',TCumul.Detail[i].GetValue('ANNEE'));
               T.AddChampSupValeur('DATEDEB',EncodeDate(AnneeI,12,31));
               T.AddChampSupValeur('DATEFIN',EncodeDate(AnneeI,12,31));
               T.AddChampSupValeur('DRACQ',TCumul.Detail[i].GetValue('ACQUIS'));
          end
          else
          begin
               T.AddChampSupValeur('LIBELLE','Cumul au '  + DateToStr(FindeMois(PlusMois(LaDate,-1))));
               T.AddChampSupValeur('ANNEE',TCumul.Detail[i].GetValue('ANNEE'));
               T.AddChampSupValeur('DATEDEB',FindeMois(PlusMois(LaDate,-1)));
               T.AddChampSupValeur('DATEFIN',FinDeMois(PlusMois(LaDate,-1)));
               T.AddChampSupValeur('DRACQ',TCumul.Detail[i].GetValue('ACQUIS'));
          end;
          T.AddChampSupValeur('FTT',0);
          T.AddChampSupValeur('FHTT',0);
          T.AddChampSupValeur('FTOTAL',0);
          T.AddChampSupValeur('DETT',0);
          T.AddChampSupValeur('DEHTT',0);
          T.AddChampSupValeur('DETOTAL',0);
          T.AddChampSupValeur('DEETAT','');
          T.AddChampSupValeur('SOLDE',0);
          T.AddChampSupValeur('SOLDETHEO',0);
          T.AddChampSupValeur('NUM','');
          T.AddChampSupValeur('STAGE','');
          T.AddChampSupValeur('MILLESIME','');
          T.AddChampSupValeur('ETAB','');
     end;
     TCumul.Free;

end;

procedure TOF_PGCOMPTEURSDIF.RemplirFormations;
var TobLesformations,T : Tob;
    i : Integer;
    Q : TQuery;
begin
     Q := OpenSQL('SELECT YEAR(PFO_DATEFIN) ANNEE,PFO_DATEDEBUT,PFO_DATEFIN,PFO_CODESTAGE,'+
     'PFO_HTPSNONTRAV,PFO_HTPSTRAV,PFO_NBREHEURE,PFO_ORDRE,PFO_MILLESIME FROM FORMATIONS '+
     'WHERE PFO_TYPEPLANPREV="DIF" AND PFO_SALARIE="'+LeSalarie+'"',True);
     TobLesformations := Tob.Create('Lesformations',Nil,-1);
     TobLesformations.LoadDetailDB('Lesformations','','',Q,False);
     Ferme(Q);
     For i := 0 To TobLesformations.Detail.Count -1 do
     begin
          T := Tob.Create('FilleEdition',TobEditDIF,-1);
          T.AddChampSupValeur('SALARIE',LeSalarie);
          T.AddChampSupValeur('TYPE','FOR');
          T.AddChampSupValeur('ORDRE',1);
          T.AddChampSupValeur('LIBELLE',RechDom('PGSTAGEFORMCOMPLET',TobLesformations.Detail[i].GetValue('PFO_CODESTAGE'),False));
          T.AddChampSupValeur('ANNEE',TobLesformations.Detail[i].GetValue('ANNEE'));
          T.AddChampSupValeur('DATEDEB',TobLesformations.Detail[i].GetValue('PFO_DATEDEBUT'));
          T.AddChampSupValeur('DATEFIN',TobLesformations.Detail[i].GetValue('PFO_DATEFIN'));
          T.AddChampSupValeur('DRACQ',0);
          T.AddChampSupValeur('DREC',0);
          T.AddChampSupValeur('FTT',TobLesformations.Detail[i].GetValue('PFO_HTPSTRAV'));
          T.AddChampSupValeur('FHTT',TobLesformations.Detail[i].GetValue('PFO_HTPSNONTRAV'));
          T.AddChampSupValeur('FTOTAL',TobLesformations.Detail[i].GetValue('PFO_NBREHEURE'));
          T.AddChampSupValeur('DETT',0);
          T.AddChampSupValeur('DEHTT',0);
          T.AddChampSupValeur('DETOTAL',0);
          T.AddChampSupValeur('DEETAT','');
          T.AddChampSupValeur('SOLDE',0);
          T.AddChampSupValeur('SOLDETHEO',0);
          T.AddChampSupValeur('NUM',TobLesformations.Detail[i].GetValue('PFO_ORDRE'));
          T.AddChampSupValeur('STAGE',TobLesformations.Detail[i].GetValue('PFO_CODESTAGE'));
          T.AddChampSupValeur('MILLESIME',TobLesformations.Detail[i].GetValue('PFO_MILLESIME'));
          T.AddChampSupValeur('ETAB','');
     end;
     TobLesformations.Free;
end;

Procedure TOF_PGCOMPTEURSDIF.RemplirDemande;
var TobLesformations,T : Tob;
    i : Integer;
    Q : TQuery;
begin
     Q := OpenSQL('SELECT PFI_CODESTAGE,PFI_MILLESIME,PFI_ETATINSCFOR,PFI_DATEDIF,'+
     'PFI_HTPSNONTRAV,PFI_HTPSTRAV,PFI_DUREESTAGE,PFI_RANG,PFI_ETABLISSEMENT FROM INSCFORMATION '+
     'WHERE (PFI_ETATINSCFOR="ATT" OR PFI_ETATINSCFOR="VAL") AND PFI_TYPEPLANPREV="DIF" AND PFI_SALARIE="'+LeSalarie+'" AND PFI_REALISE="-"',True);
     TobLesformations := Tob.Create('Lesformations',Nil,-1);
     TobLesformations.LoadDetailDB('Lesformations','','',Q,False);
     Ferme(Q);
     For i := 0 To TobLesformations.Detail.Count -1 do
     begin
          T := Tob.Create('FilleEdition',TobEditDIF,-1);
          T.AddChampSupValeur('SALARIE',LeSalarie);
          T.AddChampSupValeur('TYPE','DEMANDE');
          T.AddChampSupValeur('ORDRE',1);
          T.AddChampSupValeur('LIBELLE',RechDom('PGSTAGEFORMCOMPLET',TobLesformations.Detail[i].GetValue('PFI_CODESTAGE'),False));
          T.AddChampSupValeur('ANNEE',TobLesformations.Detail[i].GetValue('PFI_MILLESIME'));
          T.AddChampSupValeur('DATEDEB',TobLesformations.Detail[i].GetValue('PFI_DATEDIF'));
          T.AddChampSupValeur('DATEFIN',TobLesformations.Detail[i].GetValue('PFI_DATEDIF'));
          T.AddChampSupValeur('DRACQ',0);
          T.AddChampSupValeur('DREC',0);
          T.AddChampSupValeur('FTT',0);
          T.AddChampSupValeur('FHTT',0);
          T.AddChampSupValeur('FTOTAL',0);
          T.AddChampSupValeur('DETT',TobLesformations.Detail[i].GetValue('PFI_HTPSTRAV'));
          T.AddChampSupValeur('DEHTT',TobLesformations.Detail[i].GetValue('PFI_HTPSNONTRAV'));
          T.AddChampSupValeur('DETOTAL',TobLesformations.Detail[i].GetValue('PFI_DUREESTAGE'));
          T.AddChampSupValeur('DEETAT',RechDom('PGETATVALIDATION',TobLesformations.Detail[i].GetValue('PFI_ETATINSCFOR'),False));
          T.AddChampSupValeur('SOLDE',0);
          T.AddChampSupValeur('SOLDETHEO',0);
          T.AddChampSupValeur('NUM',TobLesformations.Detail[i].GetValue('PFI_RANG'));
          T.AddChampSupValeur('STAGE',TobLesformations.Detail[i].GetValue('PFI_CODESTAGE'));
          T.AddChampSupValeur('MILLESIME',TobLesformations.Detail[i].GetValue('PFI_MILLESIME'));
          T.AddChampSupValeur('ETAB',TobLesformations.Detail[i].GetValue('PFI_ETABLISSEMENT'));
     end;
     TobLesformations.Free;
end;

Procedure TOF_PGCOMPTEURSDIF.CalculerTotal;
var i : Integer;
    TotAcq,TotFTr,TotFNTr,TotF,TotDTr,TotDNTrv,TotD,Solde,SoldeEC : Double;
    aa,mm,dd : word;
    T : Tob;
begin
     TotAcq := 0;
     TotFTr := 0;
     TotFNTr := 0;
     TotF := 0;
     TotDTr := 0;
     TotDNTrv := 0;
     TotD := 0;
     For i := 0 to TobEditDIF.Detail.Count -1 do
     begin
          TotAcq := TotAcq + TobEditDIF.Detail[i].GetValue('DRACQ');
          TotFTr := TotFTr + TobEditDIF.Detail[i].GetValue('FTT');
          TotFNTr := TotFNTr + TobEditDIF.Detail[i].GetValue('FHTT');
          TotF := TotF + TobEditDIF.Detail[i].GetValue('FTOTAL');
          TotDTr := TotDTr + TobEditDIF.Detail[i].GetValue('DETT');
          TotDNTrv := TotDNTrv + TobEditDIF.Detail[i].GetValue('DEHTT');
          TotD := TotD + TobEditDIF.Detail[i].GetValue('DETOTAL');
     end;
     Solde := TotAcq - TotF;
     SoldeEC :=  TotAcq - TotF - TotD;
     T := Tob.Create('FilleEdition',TobEditDIF,-1);
     T.AddChampSupValeur('SALARIE',LeSalarie);
     T.AddChampSupValeur('TYPE','TOTAL');
     T.AddChampSupValeur('ORDRE',2);
     T.AddChampSupValeur('LIBELLE','Total au '+DateToStr(LaDate));
     decodedate(LaDate, aa, mm, dd);
     T.AddChampSupValeur('ANNEE',IntToStr(aa));
     T.AddChampSupValeur('DATEDEB',LaDate);
     T.AddChampSupValeur('DATEFIN',LaDate);
     T.AddChampSupValeur('DRACQ',TotAcq);
     T.AddChampSupValeur('FTT',TotFTr);
     T.AddChampSupValeur('FHTT',TotFNTr);
     T.AddChampSupValeur('FTOTAL',TotF);
     T.AddChampSupValeur('DETT',TotDTr);
     T.AddChampSupValeur('DEHTT',TotDNTrv);
     T.AddChampSupValeur('DETOTAL',TotD);
     T.AddChampSupValeur('DEETAT','');
     T.AddChampSupValeur('SOLDE',Solde);
     T.AddChampSupValeur('SOLDETHEO',SoldeEC);
     T.AddChampSupValeur('NUM','');
     T.AddChampSupValeur('STAGE','');
     T.AddChampSupValeur('MILLESIME','');
     T.AddChampSupValeur('ETAB','');
end;

procedure TOF_PGCOMPTEURSDIF.GrilleGetCellCanvas( ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
begin
      If ARow <> GrilleDIF.RowCount - 1 then Exit;
      Canvas.Font.Color := Clred;
      Canvas.Font.Style := [fsBold];
end;

procedure TOF_PGCOMPTEURSDIF.BtClick (Sender : TObject);
var i : Integer;
    Down : Boolean;
    HMTrad : THSystemMenu;
begin
     If TToolBarButton97(Sender).Down then Down := True
     else Down := False;
     If TToolBarButton97(Sender).Name = 'BDEMANDE' then
     begin
           For i := 1 to GrilleDIF.RowCount -1 do
           begin
                If GrilleDIF.CellValues[15,i] = 'DEMANDE' then
                begin
                     If Not Down then
                     begin
                          GrilleDIF.RowHeights[i] := -1;
                          GrilleDIF.ColWidths[8] := -1;
                          GrilleDIF.ColWidths[9] := -1;
                          GrilleDIF.ColWidths[10] := -1;
                          GrilleDIF.ColWidths[11] := -1;
                          HMtrad := THSystemMenu(GetControl('HMTRAD'));
                          HMTrad.ResizeGridColumns(GrilleDIF);
                          TToolBarButton97(Sender).Hint := 'Afficher les demandes';
                     end
                     else
                     begin
                          GrilleDIF.RowHeights[i] := 24;
                          GrilleDIF.ColWidths[8] := 60;
                          GrilleDIF.ColWidths[9] := 60;
                          GrilleDIF.ColWidths[10] := 60;
                          GrilleDIF.ColWidths[11] := 60;
                          HMTrad := THSystemMenu(GetControl('HMTRAD'));
                          HMTrad.ResizeGridColumns(GrilleDIF);
                          TToolBarButton97(Sender).Hint := 'Masquer les demandes';
                     end;
                end;
           end;
     end;
     If TToolBarButton97(Sender).Name = 'BCUMUL' then
     begin
           For i := 1 to GrilleDIF.RowCount -1 do
           begin
                If GrilleDIF.CellValues[15,i] = 'DROITS' then
                begin
                     If Not Down then
                     begin
                          GrilleDIF.RowHeights[i] := -1;
                          HMTrad := THSystemMenu(GetControl('HMTRAD'));
                          HMTrad.ResizeGridColumns(GrilleDIF);
                          TToolBarButton97(Sender).Hint := 'Afficher les cumuls';
                     end
                     else
                     begin
                          GrilleDIF.RowHeights[i] := 24;
                          HMTrad := THSystemMenu(GetControl('HMTRAD'));
                          HMTrad.ResizeGridColumns(GrilleDIF);
                          TToolBarButton97(Sender).Hint := 'Masquer les cumuls';
                     end;
                end;
           end;
     end;
     If TToolBarButton97(Sender).Name = 'BFORMATION' then
     begin
           For i := 1 to GrilleDIF.RowCount -1 do
           begin
                If GrilleDIF.CellValues[15,i] = 'FOR' then
                begin
                     If Not Down then
                     begin
                          GrilleDIF.RowHeights[i] := -1;
                          GrilleDIF.ColWidths[4] := -1;
                          GrilleDIF.ColWidths[5] := -1;
                          GrilleDIF.ColWidths[6] := -1;
                          HMTrad := THSystemMenu(GetControl('HMTRAD'));
                          HMTrad.ResizeGridColumns(GrilleDIF);
                          TToolBarButton97(Sender).Hint := 'Afficher les formations';
                     end
                     else
                     begin
                          GrilleDIF.RowHeights[i] := 24;
                          GrilleDIF.ColWidths[4] := 60;
                          GrilleDIF.ColWidths[5] := 60;
                          GrilleDIF.ColWidths[6] := 60;
                          HMTrad := THSystemMenu(GetControl('HMTRAD'));
                          HMTrad.ResizeGridColumns(GrilleDIF);
                          TToolBarButton97(Sender).Hint := 'Masquer les formations';
                     end;
                end;
           end;
     end;
end;

procedure TOF_PGCOMPTEURSDIF.GrilleDblClick(Sender : TObject);
Var Ligne : Integer;
    Rang,Stage,Millesime,Etab,St : String;
begin
        Ligne := GrilleDIF.Row;
        If GrilleDIF.CellValues[15,Ligne] = 'DEMANDE' then
        begin
             Rang := GrilleDIF.CellValues[16,Ligne];
             Stage := GrilleDIF.CellValues[17,Ligne];
             Millesime := GrilleDIF.CellValues[18,Ligne];
             Etab := GrilleDIF.CellValues[19,Ligne];
             st :=Rang + ';' + Stage + ';' + Millesime + ';' + Etab;
             AglLanceFiche('PAY','INSCFORMATION','',St,'VALIDATION;BUDGET;;;;ACTION=MODIFICATION');
        end;
        If GrilleDIF.CellValues[15,Ligne] = 'FOR' then
        begin
             Rang := GrilleDIF.CellValues[16,Ligne];
             Stage := GrilleDIF.CellValues[17,Ligne];
             Millesime := GrilleDIF.CellValues[18,Ligne];
             Etab := GrilleDIF.CellValues[19,Ligne];
             st :=Stage + ';' + Rang + ';' + Millesime + ';' + LeSalarie;
             AglLanceFiche('PAY','FORMATIONS','',St,'ACTION=CONSULTATION')
        end;
end;

Initialization
  registerclasses ( [ TOF_PGCOMPTEURSDIF ] ) ;
end.


