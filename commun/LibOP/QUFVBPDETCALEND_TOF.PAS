{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 15/06/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : QUFVBPDETCALEND ()
Mots clefs ... : TOF;QUFVBPDETCALEND
*****************************************************************}
Unit QUFVBPDETCALEND_TOF ;

Interface

Uses Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     uTob,
     forms,
     HCtrls,
     SysUtils,
     HEnt1,
     UTOF ;

Type
  TOF_QUFVBPDETCALEND = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private
     Calendrier,code:hString;
     LigneEnter,ColEnter:integer;
     OkCalendException,OkValid:boolean;
     DateDebSem:TDateTime;
     TabPrct,TabPrctModif:array [1..7] of double;
     TabCom:array [1..7] of hString;
     laTob:Tob;
     procedure RemplitGrid;
     procedure MAJTableGrid;
     procedure RemplitTabs;
     procedure BtnSolde(Sender: TObject);
     procedure BtnSuiv(Sender: TObject);
     procedure BtnPrec(Sender: TObject);
     procedure RecalculTot(F:THGrid);
     procedure CellExit(Sender: TObject; var ACol,ARow: Longint;
               var Cancel: Boolean);
     procedure CellEnter(Sender: TObject; var ACol,ARow: Longint;
               var Cancel: Boolean);
  end ;

Implementation

uses HTB97,Graphics,
     UUtil;

procedure TOF_QUFVBPDETCALEND.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPDETCALEND.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPDETCALEND.OnUpdate ;
begin
  Inherited ;
  MAJTableGrid;
  if OkValid
   then
    if MExecuteSql('UPDATE QBPDETCALENDREP SET '+code+
                   ' WHERE QBE_CALENDREP="'+calendrier+
                   '" AND QBE_DATEDEBSEM="'+USDATETIME(DateDebSem)+'"',
                   'QUFVBPDETCALEND_TOF (OnUpdate).')=0
     then MExecuteSql('INSERT INTO QBPDETCALENDREP (QBE_CALENDREP,QBE_DATEDEBSEM,'+
                      'QBE_PRCTJOUR1,QBE_PRCTJOUR2,QBE_PRCTJOUR3,QBE_PRCTJOUR4,'+
                      'QBE_PRCTJOUR5,QBE_PRCTJOUR6,QBE_PRCTJOUR7,'+
                      'QBE_COMMENTAIRE1,QBE_COMMENTAIRE2,'+
                      'QBE_COMMENTAIRE3,QBE_COMMENTAIRE4,QBE_COMMENTAIRE5,'+
                      'QBE_COMMENTAIRE6,QBE_COMMENTAIRE7) VALUES ("'+
                      Calendrier+'","'+USDATETIME(DateDebSem)+'","'+
                      STRFPOINT(TabPrctModif[1])+'","'+STRFPOINT(TabPrctModif[2])+'","'+
                      STRFPOINT(TabPrctModif[3])+'","'+STRFPOINT(TabPrctModif[4])+'","'+
                      STRFPOINT(TabPrctModif[5])+'","'+STRFPOINT(TabPrctModif[6])+'","'+
                      STRFPOINT(TabPrctModif[7])+'","'+
                      (TabCom[1])+'","'+(TabCom[2])+'","'+(TabCom[3])+'","'+
                      (TabCom[4])+'","'+(TabCom[5])+'","'+(TabCom[6])+'","'+
                      (TabCom[7])+'")',
                      'QUFVBPDETCALEND_TOF (OnUpdate).')
end ;

procedure TOF_QUFVBPDETCALEND.OnLoad ;
begin
  Inherited ;
  RemplitGrid;
  OkValid:=true;
end ;


procedure TOF_QUFVBPDETCALEND.RemplitGrid;
var i:integer;
    somme:double;   
    days: array[1..7] of hString;   
begin
 THGrid(getcontrol('GRIDLOI')).RowCount:=0;

 somme:=0;

 days[1] := TraduireMemoire('Dimanche');
 days[2] := TraduireMemoire('Lundi');
 days[3] := TraduireMemoire('Mardi');
 days[4] := TraduireMemoire('Mercredi');
 days[5] := TraduireMemoire('Jeudi');
 days[6] := TraduireMemoire('Vendredi');
 days[7] := TraduireMemoire('Samedi');

 for i:=1 to 7 do
  begin
   Latob.Detail[i-1].putValue('JOUR',days[DayOfWeek(PLUSDATE(DateDebSem,i-1,'J'))]+' '+DateTimeToStr(PLUSDATE(DateDebSem,i-1,'J')));
   Latob.Detail[i-1].putValue('PRCT',TabPrctModif[i]);
   Latob.Detail[i-1].putValue('COM',TabCom[i]);
   somme:=somme+TabPrct[i];
  end;

 laTob.PutGridDetail(THGrid(getcontrol('GRIDLOI')),true,false,'');

 SetControlText('EDTTOTAL',Format('%10.4f',[somme]));
 SetControlText('EDTSOLDE',Format('%10.4f',[100-somme]));
end;

procedure TOF_QUFVBPDETCALEND.MAJTableGrid;
var i:integer;
    com:hString;
    {$IFDEF MODE} val : hString; //EVI_TEMP_V800
    {$ELSE} val : string; {$ENDIF}
    somme:double;
begin
 code:='';
 somme:=0;

 for i := 0 to laTob.Detail.Count-1 do
  begin
   laTob.GetLigneGrid(THGrid(getcontrol('GRIDLOI')),i,'JOUR;PRCT;COM');
   val:=laTob.getValue('PRCT');           
   com:=laTob.getValue('COM');
   somme:=somme+StrToFloat(val);
   RemplaceVirguleParPoint(val);
   if code=''
    then code:=' QBE_PRCTJOUR'+IntToStr(i+1)+'="'+val+'",QBE_COMMENTAIRE'+IntToStr(i+1)+'="'+com+'" '
    else code:=code+', QBE_PRCTJOUR'+IntToStr(i+1)+'="'+val+'",QBE_COMMENTAIRE'+IntToStr(i+1)+'="'+com+'" ';
  end;


 OkValid:=false;
 if (VALEUR(Format('%10.4f',[somme]))=100)
  then OkValid:=true;
end;

procedure TOF_QUFVBPDETCALEND.RemplitTabs;
var Q:TQuery;
    i,An,NumSem:integer;
begin
 //Titre fenetre
 NumSem:=NumSemaine(DateDebSem,An);
 TForm(Ecran).caption:=TraduireMemoire('Détail du calendrier : Semaine ')+IntToStr(NumSem);
 UpdateCaption(TForm(Ecran)) ;

 OkCalendException:=false;
 THGrid(getcontrol('GRIDLOI')).ColColors[1]:=clBlack;
 //semaine exception ou semaine type
 Q:=MOPenSql('SELECT QBE_PRCTJOUR1,QBE_PRCTJOUR2,QBE_PRCTJOUR3,QBE_PRCTJOUR4,'+
              'QBE_PRCTJOUR5,QBE_PRCTJOUR6,QBE_PRCTJOUR7,'+
              'QBE_COMMENTAIRE1,QBE_COMMENTAIRE2,'+
              'QBE_COMMENTAIRE3,QBE_COMMENTAIRE4,QBE_COMMENTAIRE5,'+
              'QBE_COMMENTAIRE6,QBE_COMMENTAIRE7 '+
              ' FROM QBPDETCALENDREP '+
              'WHERE QBE_CALENDREP="'+calendrier+
              '" AND QBE_DATEDEBSEM="'+USDATETIME(DateDebSem)+'" ',
              'QUFVBPDETCALEND_TOF (RemplitTabs)',true);
  if not Q.eof
   then
    begin
     OkCalendException:=true;
     THGrid(getcontrol('GRIDLOI')).ColColors[1]:=clred;
     for i:=1 to 7 do
      begin
       TabPrct[i]:=Q.fields[i-1].asFloat;
       TabPrctModif[i]:=Q.fields[i-1].asFloat;
       TabCom[i]:=Q.fields[i-1+7].asString;
      end;
    end;
  ferme(Q);

  if not OkCalendException
   then
    begin
     //cherche semaine type du calendrier
     Q:=MOpenSql('SELECT QBN_PRCT1,QBN_PRCT2,QBN_PRCT3,QBN_PRCT4,'+
                 'QBN_PRCT5,QBN_PRCT6,QBN_PRCT7 FROM QBPCALENDREP '+
                 'WHERE QBN_CALENDREP="'+Calendrier+'"',
                 'QUFVBPDETCALEND_TOF (RemplitTabs)',true);
     if not Q.eof
      then
       begin
        for i:=1 to 7 do
         begin
          TabPrct[i]:=Q.fields[i-1].asFloat;
          TabPrctModif[i]:=Q.fields[i-1].asFloat;
          TabCom[i]:='';
         end;
       end;
     ferme(Q);
    end;
end;


procedure TOF_QUFVBPDETCALEND.OnArgument (S : String ) ;
var i,An,NumSem:integer;
     Tf: Tob;
begin
  Inherited ;
  Calendrier:=TrouveArgument(S,'CALEND','');
  DateDebSem:=StrToDateTime(TrouveArgument(S,'DELAI',''));
  NumSem:=NumSemaine(DateDebSem,An);
  DateDebSem:=PremierJourSemaine(NumSem,An);

  THGrid(getcontrol('GRIDLOI')).ColFormats[1]:='#,####0.0000';
  THGrid(getcontrol('GRIDLOI')).ColAligns[1]:=taRightJustify;
  THGrid(getcontrol('GRIDLOI')).ColTypes[1]:='F';
 // THGrid(getcontrol('GRIDLOI')).ColFormats[2]:='CB=QUTBPCALENDCOM';
  THGrid(getcontrol('GRIDLOI')).OnCellExit:=CellExit;
  THGrid(getcontrol('GRIDLOI')).OnCellEnter:=CellEnter;

  TToolBarButton97(getcontrol('BTNSOLDE')).OnClick := BtnSolde;     
  TToolBarButton97(getcontrol('BTNPREC')).OnClick := BtnPrec;
  TToolBarButton97(getcontrol('BTNSUIV')).OnClick := BtnSuiv;

  //création de la tob
  laTob := Tob.Create('_QT', nil, -1);
  for i:=0 to 6 do
   begin
    Tf := Tob.Create('_QGRF', latob, -1);
    Tf.AddChampSupValeur('JOUR', '');    
    Tf.AddChampSupValeur('PRCT', '');
    Tf.AddChampSupValeur('COM', '');
   end;
  RemplitTabs;
end ;

procedure TOF_QUFVBPDETCALEND.OnClose ;
begin
  Inherited ;
  LaTob.free;
  
  if not OkValid
   then LastError:=1
   else LastError:=0;
  LastErrorMsg:=TraduireMemoire('La somme des pourcentages doit être égale à 100.');
end ;

procedure TOF_QUFVBPDETCALEND.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPDETCALEND.OnCancel () ;
begin
  Inherited ;
end ;


procedure TOF_QUFVBPDETCALEND.CellEnter(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
var G:THGrid ;
begin
 G:=THGrid(Sender) ;
 LigneEnter:=G.row;
 ColEnter:=G.col;
end;

procedure TOF_QUFVBPDETCALEND.RecalculTot(F:THGrid);
var i:Integer ;
    somme:double ;
begin
 somme:=0 ;
 for i:=0 to F.RowCount-1 do
  begin
   somme:=somme+Valeur(F.Cells[1,i]) ;
   TabPrctModif[i+1]:=Valeur(F.Cells[1,i]);
   TabCom[i+1]:=(F.CellValues[2,i]);
  end;
 SetControlText('EDTTOTAL',Format('%10.4f',[somme]));
 SetControlText('EDTSOLDE',Format('%10.4f',[100-somme]));
end ;

procedure TOF_QUFVBPDETCALEND.CellExit(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
var G:THGrid ;
begin
 G:=THGrid(Sender) ;
 RecalculTot(G) ;
end;


procedure TOF_QUFVBPDETCALEND.BtnSolde(Sender: TObject);
var val,somme:double;
    i:integer;
begin
 MAJTableGrid;
 laTob.GetLigneGrid(THGrid(getcontrol('GRIDLOI')),LigneEnter,'JOUR;PRCT;COM');
 val:=laTob.getValue('PRCT');
 laTob.SetDouble('PRCT',val+StrToFloat(GetControlText('EDTSOLDE')));
 laTob.PutLigneGrid(THGrid(getcontrol('GRIDLOI')),LigneEnter,false,false,'JOUR;PRCT;COM');
 somme:=0;
 for i := 0 to laTob.Detail.Count-1 do
  begin
   laTob.GetLigneGrid(THGrid(getcontrol('GRIDLOI')),i,'JOUR;PRCT;COM');
   val:=laTob.getValue('PRCT');
   somme:=somme+val;
  end;

 SetControlText('EDTTOTAL',Format('%10.4f',[somme]));
 SetControlText('EDTSOLDE',Format('%10.4f',[100-somme]));
end;

procedure TOF_QUFVBPDETCALEND.BtnPrec(Sender: TObject);
begin
 DateDebSem:=DateDebSem-7;
 RemplitTabs;
 RemplitGrid;
end;

procedure TOF_QUFVBPDETCALEND.BtnSuiv(Sender: TObject);
begin
 DateDebSem:=DateDebSem+7;
 RemplitTabs;
 RemplitGrid;
end;

Initialization
  registerclasses ( [ TOF_QUFVBPDETCALEND ] ) ; 
end.
