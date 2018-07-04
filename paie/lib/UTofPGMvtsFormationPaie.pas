{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 28/04/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMVTSFORMPAIE ()
Mots clefs ... : TOF;PGMVTSFORMPAIE
*****************************************************************}
Unit UTofPGMvtsFormationPaie ;

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
     forms,
     uTob,  
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1,
     HSysMenu,
     HMsgBox, 
     UTOF ; 

Type
  TOF_PGMVTSFORMPAIE = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    private
    GMvts : THGrid;
    LeStage,NumOrdre,LeMillesime : String;
    procedure AlimenteGrille;
  end ;

Implementation

procedure TOF_PGMVTSFORMPAIE.OnArgument (S : String ) ;
begin
  Inherited ;
  LeStage := ReadTokenPipe(S,';');
  NumOrdre := ReadTokenPipe(S,';');
  LeMillesime := ReadTokenPipe(S,';');
  GMvts := THGrid (GetControl('GMOUVEMENTS'));
  GMvts.ColFormats[3] := ShortDateFormat;
  GMvts.ColFormats[4] := ShortDateFormat;
  GMvts.ColFormats[5] := '# ##0.00';
  GMvts.ColFormats[6] := 'CB=PGOUINONGRAPHIQUE';
  GMvts.ColDrawingModes[6]:= 'IMAGE';
  Ecran.Caption := 'Mouvements générés pour la formation '+RechDom('PGSTAGEFORM',LeStage,False);
  AlimenteGrille;
end ;

procedure TOF_PGMVTSFORMPAIE.AlimenteGrille;
var Q : TQuery;
    TobGrille,TG : Tob;
    TobStagiaire,TobMvtDIF : Tob;
    i,a,b : Integer;
    Salarie,FJ,DJ,TypeFor : String;
    DateD,DateF : TDateTime;
    AncDateD,AncDateF : TDateTime;
    NbhT,NbHNonT,NbHeures : Double;
    Travail,HorsTravail : String;
    Rubrique,TypeConge,TypeAlim : String;
    HeuresTAlim,HeuresNonTAlim : Double;
    RubriqueCreer : Boolean;
    St,StChampGrid : String;
    HMTRAD : THSystemMenu;
    DDAbs,DFAbs : TDateTime;
    TobAbs : Tob;
    LesHeuresCreer : Double;
    Predef : String;
    DateDRub,DateFRub : TDateTime;
begin
     TobGrille := Tob.Create('AffichageMvt',Nil,-1);
     Q := OpenSQL('SELECT PSS_DATEDEBUT,PSS_DATEFIN,PSS_DEBUTDJ,PSS_FINDJ FROM SESSIONSTAGE WHERE '+
     'PSS_CODESTAGE="'+LeStage+'" AND '+
     'PSS_ORDRE='+NumOrdre+' AND '+
     'PSS_MILLESIME="'+LeMillesime+'"',True);
     If Not Q.Eof then
     begin
          DJ := Q.FindField('PSS_DEBUTDJ').AsString;
          FJ := Q.FindField('PSS_FINDJ').AsString;
     end;
     Ferme(Q);
     Q := OpenSQL('SELECT * FROM FORMATIONS WHERE '+
     'PFO_CODESTAGE="'+LeStage+'" AND '+
     'PFO_ORDRE='+NumOrdre+' AND '+
     'PFO_MILLESIME="'+LeMillesime+'" AND '+
      'PFO_EFFECTUE="X"',True);
     TobStagiaire :=Tob.Create('LesDif',Nil,-1);
     TobStagiaire.LoadDetailDB('LesDif','','',Q,False);
     ferme(Q);
     For i := 0 to TobStagiaire.Detail.Count - 1 do
     begin
          Salarie := TobStagiaire.Detail[i].GetValue('PFO_SALARIE');
          TypeFor := TobStagiaire.Detail[i].GetValue('PFO_TYPEPLANPREV');
          NbhT := TobStagiaire.Detail[i].GetValue('PFO_HTPSTRAV');
          NbhNonT := TobStagiaire.Detail[i].GetValue('PFO_HTPSNONTRAV');
          DateD := TobStagiaire.Detail[i].GetValue('PFO_DATEDEBUT');
          DateF := TobStagiaire.Detail[i].GetValue('PFO_DATEFIN');
          DateDRub := DebutDeMois(TobStagiaire.Detail[i].GetValue('PFO_DATEPAIE'));
          DateFRub := FinDemois(TobStagiaire.Detail[i].GetValue('PFO_DATEPAIE'));
          Q := OpenSQL('SELECT * FROM PARAMFORMABS WHERE ##PPF_PREDEFINI## PPF_TYPEPLANPREV="'+TypeFor+'"',True);
          TobMvtDIF := Tob.Create('Lesmvts',Nil,-1);
          TobMvtDIF.LoadDetailDB('Lesmvts','','',Q,False);
          Ferme(Q);
          For a := 0 TO TobMvtDIF.Detail.Count - 1 do
          begin
               Travail := TobMvtDIF.Detail[a].GetValue('PPF_CHEURETRAV');
               HorsTravail := TobMvtDIF.Detail[a].GetValue('PPF_CHEURENONTRAV');
               Predef := TobMvtDIF.Detail[a].GetValue('PPF_PREDEFINI');
               If Predef <> 'DOS' then
               begin
                     If TobMvtDIF.FindFirst(['PPF_PREDEFINI','PPF_CHEURETRAV','PPF_CHEURENONTRAV','PPF_TYPEPLANPREV'],['DOS',Travail,HorsTravail,TypeFor],False) <> Nil then Continue;
                     If Predef = 'CEG' then
                     begin
                          If TobMvtDIF.FindFirst(['PPF_PREDEFINI','PPF_CHEURETRAV','PPF_CHEURENONTRAV','PPF_TYPEPLANPREV'],['STD',Travail,HorsTravail,TypeFor],False) <> Nil then Continue;
                     end;
               end;
               HeuresTAlim := 0;
               HeuresNonTAlim := 0;
               If Travail = 'X' then HeuresTAlim := NbhT;
               If HorsTravail = 'X' then HeuresNonTAlim := NbhNonT;
               If TobMvtDIF.Detail[a].GetValue('PPF_ALIMABS') = 'X' then
               begin
                     TypeConge := TobMvtDIF.Detail[a].GetValue('PPF_TYPECONGE');
                     Q := OpenSQL('SELECT * FROM ABSENCESALARIE WHERE '+
                     'PCN_SALARIE="'+Salarie+'" AND '+
                     'PCN_DATEDEBUTABS>="'+UsDateTime(DateD)+'" AND '+
                     'PCN_DATEFINABS<="'+UsDateTime(DateF)+'" AND PCN_TYPECONGE="'+TypeConge+'"',True);
                     TobAbs := Tob.Create('LesAbsences',Nil,-1);
                     TobAbs.LoadDetailDB('LesAbsences','','',Q,False);
                     LesHeuresCreer := 0;
                     For b := 0 to TobAbs.Detail.Count - 1 do
                     begin
                          DDAbs := TobAbs.Detail[b].GetValue('PCN_DATEDEBUTABS');
                          DFAbs := TobAbs.Detail[b].GetValue('PCN_DATEFINABS');
                          NbHeures := TobAbs.Detail[b].GetValue('PCN_HEURES');
                          LesHeuresCreer := LesHeuresCreer + NbHeures;
                          TG := Tob.Create('LigneGRILLE',TobGrille,-1);
                          TG.AddChampSupValeur('SALARIE',Salarie);
                          TG.AddChampSupValeur('NOM',RechDom('PGSALARIE',Salarie,False));
                          TG.AddChampSupValeur('TYPE','ABSENCE');
                          TG.AddChampSupValeur('DATEDEB',DDAbs);
                          TG.AddChampSupValeur('DATEFIN',DFAbs);
                          TG.AddChampSupValeur('NBHEURES',NbHeures);
                          TG.AddChampSupValeur('ETAT','OUI');

                     end;
                     TobAbs.Free;
                     Ferme(Q);
                     If LesHeuresCreer < (HeuresNonTAlim + HeuresTAlim) then
                     begin
                          st := 'SELECT PCN_DATEDEBUTABS,PCN_DATEFINABS,PCN_TYPECONGE,PCN_TYPEMVT FROM ABSENCESALARIE ' +
                          ' WHERE PCN_SALARIE = "' + Salarie + '" AND PCN_DEBUTDJ="' + DJ + '" AND PCN_FINDJ="' + FJ + '"' +
                          ' AND PCN_TYPECONGE <> "'+TypeConge+'" ' +
                          ' AND ((PCN_DATEDEBUTABS >="' + usdatetime(DateD) + '" AND PCN_DATEDEBUTABS <= "' +
                          usdatetime(DateF) + '")' +
                          'OR (PCN_DATEFINABS >="' + usdatetime(DateD) + '" AND PCN_DATEDEBUTABS <= "' +
                          usdatetime(DateF) + '")' +
                          'OR (PCN_DATEFINABS >="' + usdatetime(DateF) + '" AND PCN_DATEDEBUTABS <= "' +
                          usdatetime(DateD) + '")' +
                          'OR(PCN_DATEFINABS <="' + usdatetime(DateF) + '" AND PCN_DATEDEBUTABS >= "' +
                          usdatetime(DateD) + '"))';
                          If ExisteSQL(St) then
                          begin
                               TG := Tob.Create('LigneGRILLE',TobGrille,-1);
                               TG.AddChampSupValeur('SALARIE',Salarie);
                               TG.AddChampSupValeur('NOM',RechDom('PGSALARIE',Salarie,False));
                               TG.AddChampSupValeur('TYPE','ABSENCE');
                               TG.AddChampSupValeur('DATEDEB',DateD);
                               TG.AddChampSupValeur('DATEFIN',DateF);
                               TG.AddChampSupValeur('NBHEURES',LesHeuresCreer - (HeuresTAlim+HeuresNonTAlim));
                               TG.AddChampSupValeur('ETAT','NON');
                          end
                          else
                          begin
                               TG := Tob.Create('LigneGRILLE',TobGrille,-1);
                               TG.AddChampSupValeur('SALARIE',Salarie);
                               TG.AddChampSupValeur('NOM',RechDom('PGSALARIE',Salarie,False));
                               TG.AddChampSupValeur('TYPE','ABSENCE');
                               TG.AddChampSupValeur('DATEDEB',DateD);
                               TG.AddChampSupValeur('DATEFIN',DateF);
                               TG.AddChampSupValeur('NBHEURES',HeuresTAlim+HeuresNonTAlim);
                               TG.AddChampSupValeur('ETAT','NON');
                          end;
                     end;
               end;
               If TobMvtDIF.Detail[a].GetValue('PPF_ALIMRUB') = 'X' then
               begin
                    Rubrique := TobMvtDIF.Detail[a].GetValue('PPF_RUBRIQUE');
                    Travail := TobMvtDIF.Detail[a].GetValue('PPF_CHEURETRAV');
                    HorsTravail := TobMvtDIF.Detail[a].GetValue('PPF_CHEURENONTRAV');
                    TypeAlim := TobMvtDIF.Detail[a].GetValue('PPF_ALIMENT');
                    RubriqueCreer := False;
                    HeuresTAlim := 0;
                    HeuresNonTAlim := 0;
                    If Travail = 'X' then HeuresTAlim := NbhT;
                    If HorsTravail = 'X' then HeuresNonTAlim := NbhNonT;
                    Q := OpenSQL('SELECT * FROM HISTOSAISRUB WHERE '+
                    'PSD_ORIGINEMVT="MLB" AND '+
                    'PSD_SALARIE="'+Salarie+'" AND '+
                    'PSD_DATEDEBUT="'+UsDateTime(DateDRub)+'" AND '+
                    'PSD_DATEFIN="'+UsDateTime(DateFRub)+'" AND '+
                    'PSD_RUBRIQUE="'+Rubrique+'" AND '+
                    'PSD_ORDRE=200',True);
                    If Not Q.Eof then
                    begin
                         RubriqueCreer := True;
                         TG := Tob.Create('LigneGRILLE',TobGrille,-1);
                         TG.AddChampSupValeur('SALARIE',Salarie);
                         TG.AddChampSupValeur('NOM',RechDom('PGSALARIE',Salarie,False));
                         TG.AddChampSupValeur('TYPE','RUBRIQUE');
                         TG.AddChampSupValeur('DATEDEB',DateDRub);
                         TG.AddChampSupValeur('DATEFIN',DateFRub);
                         TG.AddChampSupValeur('NBHEURES',HeuresTAlim+HeuresNonTAlim);
                         TG.AddChampSupValeur('ETAT','OUI');
                    end;
                    ferme(Q);
                    If (Not RubriqueCreer) and ((HeuresTAlim + HeuresNonTAlim) > 0) then
                    begin
                         TG := Tob.Create('LigneGRILLE',TobGrille,-1);
                         TG.AddChampSupValeur('SALARIE',Salarie);
                         TG.AddChampSupValeur('NOM',RechDom('PGSALARIE',Salarie,False));
                         TG.AddChampSupValeur('TYPE','RUBRIQUE');
                         TG.AddChampSupValeur('DATEDEB',DateDRub);
                         TG.AddChampSupValeur('DATEFIN',DateFRub);
                         TG.AddChampSupValeur('NBHEURES',HeuresTAlim+HeuresNonTAlim);
                         TG.AddChampSupValeur('ETAT','OUI');
                    end;
               end;
          end;
          TobMvtDIF.Free;
     end;
     TobStagiaire.Free;
     StChampGrid := 'SALARIE;NOM;TYPE;DATEDEB;DATEFIN;NBHEURES;ETAT';
     If TobGrille.Detail.Count = 0 then GMvts.RowCount := 2
     else GMvts.RowCount := TobGrille.Detail.Count + 1;
     TobGrille.PutGridDetail(GMvts,False,False,StChampGrid,False);
     HMTrad.ResizeGridColumns(GMvts);
     TobGrille.Free;
end;


Initialization
  registerclasses ( [ TOF_PGMVTSFORMPAIE ] ) ;
end.

