{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 30/06/2016
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTARTDISPO ()
Mots clefs ... : TOF;BTARTDISPO
*****************************************************************}
Unit BTARTDISPO_TOF ;

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
     uTob, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     UTOF ;

Type
  TOF_BTARTDISPO = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    ThisGuid : String;
    //
    CRUPTURE : TRadioButton;
    CSURSTOCK: TRadioButton;

    procedure CalculeStockDate (DateFin : Tdatetime);
    procedure ControleRupture;
    procedure ControleSurStock;
    function  GetInfoInventaire (CodeDepot,Article : string; DateFin : TdateTime;var DateDep : TDateTime; var QteIni,PMAPINI : double) : boolean;
    procedure NettoieTable;

  end ;

Implementation
uses UtilPmapCalcul,galPatience;

procedure TOF_BTARTDISPO.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTARTDISPO.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTARTDISPO.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTARTDISPO.OnLoad ;
var SQL,SQLPlus : String;
begin
  Inherited ;

  NettoieTable;

  Sql := 'INSERT INTO BDISPODATE (BSK_GUID,BSK_DEPOT,BSK_ARTICLE,BSK_PHYSIQUE,BSK_PMAP,BSK_EMPLACEMENT,BSK_CALCULABLE) '+
         'SELECT "'+ThisGuid+'",GQ_DEPOT,GQ_ARTICLE,GQ_PHYSIQUE,GQ_PMAP,GQ_EMPLACEMENT,"X" FROM DISPO '+
         'LEFT JOIN ARTICLE ON GA_ARTICLE=GQ_ARTICLE '+
         'WHERE ';
  SQLPlus := '';
  if (THEdit(getControl('GA_CODEARTICLE')).Text <> '') then
  begin
    SqlPlus := SqlPlus + 'GA_CODEARTICLE >= "'+THEdit(getControl('GA_CODEARTICLE')).Text+'"';
  end;
  if (THEdit(getControl('GA_CODEARTICLE_')).Text <> '') then
  begin
    if SqlPlus <> '' then SQlPlus := SQLPLus + ' AND ';
    SqlPlus := SqlPlus + 'GA_CODEARTICLE <= "'+THEdit(getControl('GA_CODEARTICLE_')).Text+'"';
  end;
  if THEdit(getControl('GA_FOURNPRINC')).Text <> '' then
  begin
    if SqlPlus <> '' then SQlPlus := SQLPLus + ' AND ';
    SqlPlus := SqlPlus + 'GA_FOURNPRINC >= "'+THEdit(getControl('GA_FOURNPRINC')).Text+'"';
  end;
  if THEdit(getControl('GA_FOURNPRINC_')).Text <> '' then
  begin
    if SqlPlus <> '' then SQlPlus := SQLPLus + ' AND ';
    SqlPlus := SqlPlus + 'GA_FOURNPRINC <= "'+THEdit(getControl('GA_FOURNPRINC_')).Text+'"';
  end;
  if THValComboBox(getControl('BSK_DEPOT')).Value <> '' then
  begin
    if SQlPLus <> '' then SQLPlus := SQLPlus + ' AND ';
    SQLPlus := SQLPLus + 'GQ_DEPOT = "'+THValComboBox(getControl('BSK_DEPOT')).Value+'" ';
  end;
  if THEdit(getControl('GQ_EMPLACEMENT')).Text <> '' then
  begin
    if SQlPLus <> '' then SQLPlus := SQLPlus + ' AND ';
    SQLPlus := SQLPLus + 'GQ_EMPLACEMENT = "'+THEdit(getControl('GQ_EMPLACEMENT')).Text+'"';
  end;
  SQl := SQL + SQlplus;
  ExecuteSql(SQL);
  //
  if StrToDate(Thedit(getControl('DATEFIN')).text)=StrToDate(DateToStr(Now)) then
  begin
    if CRUPTURE.Checked then
      ControleRupture
    else if CSURSTOCK.Checked then
      ControleSurStock
    else Exit;
  end
  else
    CalculeStockDate (StrToDate(Thedit(getControl('DATEFIN')).text));
  //
  ExecuteSQL ('DELETE FROM BDISPODATE WHERE BSK_CALCULABLE="-"');
  //
end ;

//FV1 : 21/06/2017 - FS#2601 - VIROT : Erreur si édition Etat de stock à date et option cochée "En surstock"
Procedure TOF_BTARTDISPO.ControleSurStock;
var XX      : TFPatience;
    QQ      : TQuery;
    Article : string;
    Depot   : string;
Begin

  XX := FenetrePatience('Contrôle des Article en sur Stock au ' + DateToStr(Now), aoMilieu, False,true);
  XX.lAide.Caption := 'Traitement des données...';
  XX.lcreation.visible := false ;
  XX.StartK2000 ;
  XX.Refresh;

  TRY
    QQ := OpenSql ('SELECT * FROM BDISPODATE WHERE BSK_GUID="'+ThisGuid+'"',true,-1,'',true);
    if not QQ.eof then
    begin
      QQ.first;
      repeat
        Article   := QQ.FindField('BSK_ARTICLE').AsString;
        Depot     := QQ.FindField('BSK_DEPOT').AsString;
        if not ExisteSQL('SELECT GQ_ARTICLE FROM DISPO WHERE GQ_DEPOT = "' + Depot + '" AND GQ_ARTICLE = "' + Article + '" AND ((GQ_PHYSIQUE - GQ_RESERVECLI - GQ_PREPACLI + GQ_RESERVEFOU) > GQ_STOCKMAX)') then
        begin
          //Suppression de l'enregistrement DSK
          ExecuteSql ('UPDATE BDISPODATE '+
          'SET '+
          'BSK_CALCULABLE="-" '+
          'WHERE '+
          'BSK_GUID="'+ThisGuid+' " AND '+
          'BSK_ARTICLE="'+Article+'" AND '+
          'BSK_DEPOT="'+Depot+'"');
        end;
        QQ.next;
      until QQ.eof;
    end;
  finally
    XX.StopK2000 ;
    XX.free;
  end;

end;

Procedure TOF_BTARTDISPO.ControleRupture;
var XX      : TFPatience;
    QQ      : TQuery;
    Article : string;
    Depot   : string;
Begin

  XX := FenetrePatience('Contrôle des Article en rupture au ' + DateToStr(Now), aoMilieu, False,true);
  XX.lAide.Caption := 'Traitement des données...';
  XX.lcreation.visible := false ;
  XX.StartK2000 ;
  XX.Refresh;

  TRY
    QQ := OpenSql ('SELECT * FROM BDISPODATE WHERE BSK_GUID="'+ThisGuid+'"',true,-1,'',true);
    if not QQ.eof then
    begin
      QQ.first;
      repeat
        Article   := QQ.FindField('BSK_ARTICLE').AsString;
        Depot     := QQ.FindField('BSK_DEPOT').AsString;
        if not ExisteSQL('SELECT GQ_ARTICLE FROM DISPO WHERE GQ_DEPOT = "' + Depot + '" AND GQ_ARTICLE = "' + Article + '" AND ((GQ_PHYSIQUE - GQ_RESERVECLI - GQ_PREPACLI + GQ_RESERVEFOU) < GQ_STOCKMIN)') then
        begin
          //Suppression de l'enregistrement DSK
          ExecuteSql ('UPDATE BDISPODATE '+
          'SET '+
          'BSK_CALCULABLE="-" '+
          'WHERE '+
          'BSK_GUID="'+ThisGuid+' " AND '+
          'BSK_ARTICLE="'+Article+'" AND '+
          'BSK_DEPOT="'+Depot+'"');
        end;
        QQ.next;
      until QQ.eof;
    end;
  finally
    XX.StopK2000 ;
    XX.free;
  end;

end;


procedure TOF_BTARTDISPO.OnArgument (S : String ) ;
begin
  Inherited ;
  ThisGuid := AglGetGuid();
  THEdit(GetControl('BSK_GUID')).Text := ThisGuid;
  NettoieTable;

  CRUPTURE := TRadioButton(GetControl('CRUPTURE'));
  CSURSTOCK:= TRadioButton(GetControl('CSURSTOCK'));

end ;

procedure TOF_BTARTDISPO.OnClose ;
begin
  NettoieTable;
  Inherited ;
end ;

procedure TOF_BTARTDISPO.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTARTDISPO.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTARTDISPO.NettoieTable;
begin
  ExecuteSql ('DELETE FROM BDISPODATE WHERE BSK_GUID="'+ThisGuid+'"');
end;

procedure TOF_BTARTDISPO.CalculeStockDate(DateFin: Tdatetime);
var XX : TFPatience;
    QQ : TQuery;
    TheDateDep : TdateTime;
    QteIni,PmapIni,PmapStock,QteStock : double;
    fNomTable,Article,Depot : string;
begin

  fNomTable := ConstitueNomTemp;

  XX := FenetrePatience('Constitution du stock en date du '+DateToStr(Datefin),aoMilieu, False,true);
  XX.lAide.Caption := 'Traitement des données...';
  XX.lcreation.visible := false ;
  XX.StartK2000 ;
  XX.Refresh;

  TRY
    QQ := OpenSql ('SELECT * FROM BDISPODATE WHERE BSK_GUID="'+ThisGuid+'"',true,-1,'',true);
    if not QQ.eof then
    begin
      QQ.first;
      repeat
        PmapStock := 0;
        QteStock  := 0;
        Article   := QQ.FindField('BSK_ARTICLE').AsString;
        Depot     := QQ.FindField('BSK_DEPOT').AsString;
        //FV1 : 21/06/2017 - FS#2601 - VIROT : Erreur si édition Etat de stock à date et option cochée "En surstock"
        //Si rupture de stock
        if CRUPTURE.Checked then
        begin
          if not ExisteSQL('SELECT GQ_ARTICLE FROM DISPO WHERE GQ_DEPOT = "' + Depot + '" AND GQ_ARTICLE = "' + Article + '" AND ((GQ_PHYSIQUE - GQ_RESERVECLI - GQ_PREPACLI + GQ_RESERVEFOU) < GQ_STOCKMIN)') then
          begin
            //Suppression de l'enregistrement DSK
            ExecuteSql ('UPDATE BDISPODATE '+
            'SET '+
            'BSK_CALCULABLE="-" '+
            'WHERE '+
            'BSK_GUID="'+ThisGuid+' " AND '+
            'BSK_ARTICLE="'+Article+'" AND '+
            'BSK_DEPOT="'+Depot+'"');
            QQ.next;
            continue;
          end;
        end
        //FV1 : 21/06/2017 - FS#2601 - VIROT : Erreur si édition Etat de stock à date et option cochée "En surstock"
        else If CSURSTOCK.Checked then //Si en surStock
        begin
          if not ExisteSQL('SELECT GQ_ARTICLE FROM DISPO WHERE GQ_DEPOT = "' + Depot + '" AND GQ_ARTICLE = "' + Article + '" AND ((GQ_PHYSIQUE - GQ_RESERVECLI - GQ_PREPACLI + GQ_RESERVEFOU) > GQ_STOCKMAX)') then
          begin
            //Suppression de l'enregistrement DSK
            ExecuteSql ('UPDATE BDISPODATE '+
            'SET '+
            'BSK_CALCULABLE="-" '+
            'WHERE '+
            'BSK_GUID="'+ThisGuid+' " AND '+
            'BSK_ARTICLE="'+Article+'" AND '+
            'BSK_DEPOT="'+Depot+'"');
            QQ.next;
            continue;
          end;
        end;

        if GetInfoInventaire (Depot,Article,DateFin,TheDateDep,QteIni,PMAPINI) then
        begin
          DropTableExist (fNomtable);
          ConstitueTableTemp(fNomTable,Article,Depot,TheDateDep,DateFin);
          //
          if CalculsTableTempo(fNomTable,Depot, Article,PmapStock,QteStock,PMAPIni,QteIni) then
          begin
            ExecuteSql ('UPDATE BDISPODATE '+
                        'SET '+
                        'BSK_PHYSIQUE= '+strfpoint(QteStock)+','+
                        'BSK_PMAP= '+strfpoint(PmapStock)+' '+
                        'WHERE '+
                        'BSK_GUID="'+ThisGuid+' " AND '+
                        'BSK_ARTICLE="'+Article+'" AND '+
                        'BSK_DEPOT="'+Depot+'"');
          end else
          begin
            ExecuteSql ('UPDATE BDISPODATE '+
                        'SET '+
                        'BSK_PHYSIQUE= '+strfpoint(QteIni)+','+
                        'BSK_PMAP= '+strfpoint(PMAPINI)+' '+
                        'WHERE '+
                        'BSK_GUID="'+ThisGuid+' " AND '+
                        'BSK_ARTICLE="'+Article+'" AND '+
                        'BSK_DEPOT="'+Depot+'"');
          end;
        end else
        begin
          ExecuteSql ('UPDATE BDISPODATE '+
                      'SET '+
                      'BSK_CALCULABLE="-" '+
                      'WHERE '+
                      'BSK_GUID="'+ThisGuid+' " AND '+
                      'BSK_ARTICLE="'+Article+'" AND '+
                      'BSK_DEPOT="'+Depot+'"');
        end;
        QQ.next;
      until QQ.eof;
    end;
  FINALLY
    XX.StopK2000 ;
    XX.free;
  END;

end;

function TOF_BTARTDISPO.GetInfoInventaire(CodeDepot, Article: string;
  DateFin: TdateTime; var DateDep: TDateTime; var QteIni,
  PMAPINI: double): boolean;
var QQ: TQuery;
    Okok : boolean;
begin
  result := false;
  Okok := false;
  QQ := OpenSql ('SELECT * FROM DISPO WHERE '+
                 'GQ_DEPOT="'+CodeDepot+'" AND '+
                 'GQ_ARTICLE="'+Article+'"',true,1,'',true);
  if not QQ.eof then
  begin
    if DateFin >= QQ.FindField('GQ_DATEINV').AsDateTime then
    begin
      okok    := true;
      result  := true;
      DateDep := QQ.findField('GQ_DATEINV').AsDateTime;
      QteIni  := QQ.findField('GQ_STOCKINV').AsFloat;
      PMAPINI := QQ.findField('GQ_PRIXINV').AsFloat;
    end;
  end;
  ferme (QQ);
  if not okok then
  begin
    QQ := OpenSql ('SELECT ##TOP 1## * FROM BHISTOINV WHERE '+
                   'BHI_DEPOT="'+CodeDepot+'" AND '+
                   'BHI_ARTICLE="'+Article+'" AND '+
                   'BHI_DATEINV <= "'+USDATETIME(DateFin)+'" '+
                   'ORDER BY BHI_DATEINV DESC',true,1,'',true);
    result := not QQ.eof;
    if not QQ.eof then
    begin
      DateDep := QQ.findField('BHI_DATEINV').AsDateTime;
      QteIni := QQ.findField('BHI_STOCKINV').AsFloat;
      PMAPINI := QQ.findField('BHI_PRIXINV').AsFloat;
    end;
  end;
  ferme (QQ);
end;

Initialization
  registerclasses ( [ TOF_BTARTDISPO ] ) ;
end.
