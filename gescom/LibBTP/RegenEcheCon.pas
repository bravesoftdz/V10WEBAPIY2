unit RegenEcheCon;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, Hgauge, StdCtrls, Hctrls,UTOB,ParamSoc,hmsgbox,Ent1,HEnt1,
{$IFDEF EAGLCLIENT}
  maineagl
{$ELSE}
  DBCtrls, Db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} fe_main
{$ENDIF}
  ;

type
  TFREGENECHECON = class(TForm)
    Titre: THLabel;
    PB: TEnhancedGauge;
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
    procedure LanceTrait;
  end;

procedure RegenereEcheContrats;

implementation

{$R *.dfm}

procedure RegenereEcheContrats;
var XX : TFREGENECHECON;
begin
	XX := TFREGENECHECON.Create(application);
  TRY
  	XX.Show;
    XX.lanceTrait;
  FINALLY
  	XX.free;
  END;
end;

procedure TFREGENECHECON.LanceTrait;
var TOBF,TOBL : TOB;
		QQ : TQuery;
    termeEcheDef,termeEche,periodEcheDef,periodEche : string;
    InterValEcheDef,InterValEche,Indice : integer;
    Date1 : TdateTime;
begin
  Titre.Caption := 'Traitement en cours ...';
  termeEcheDef := getparamSocSecur('SO_AFTERMEECHE','ACH');
  periodEcheDef := getparamSocSecur('SO_AFPERIODICTE','M');
  InterValEcheDef := getparamSocSecur('SO_AFINTERVAL',1);
  //
  Titre.visible := true;
  Titre.Refresh;
  TOBF := TOB.Create ('LES ECHEANCES',nil,-1);
  TRY
    QQ := OpenSql ('SELECT FACTAFF.*,AFF_TERMEECHEANCE,AFF_PERIODICITE,AFF_INTERVALGENER,AFF_METHECHEANCE FROM FACTAFF '+
    							 'LEFT JOIN AFFAIRE ON AFF_AFFAIRE=FACTAFF.AFA_AFFAIRE '+
                   'WHERE FACTAFF.AFA_DATEDEBUTFAC='+DateToStr(IDate1900)+ ' OR FACTAFF.AFA_DATEDEBUTFAC IS NULL',true,-1,'',true) ;
    if not QQ.eof then
    begin
    	TOBF.LoadDetailDB('FACTAFF','','',QQ,false);
			PB.Visible := true;
      PB.MinValue := 0;
      PB.MaxValue  := TOBF.detail.count -1;
      PB.Progress := 0;
      for indice := 0 to TOBF.detail.count -1 do
      begin
        TOBL := TOBF.detail[Indice];
        if TOBL.GetValue('AFF_TERMEECHEANCE')='' then
        begin
        	termeEche := termeEcheDef;
        end else
        begin
        	termeEche := TOBL.getValue('AFF_TERMEECHEANCE');
        end;
        if TOBL.GetValue('AFF_PERIODICITE')='' then
        begin
        	periodEche := periodEcheDef;
        end else
        begin
        	periodEche := TOBL.getValue('AFF_PERIODICITE');
        end;
        if TOBL.GetValue('AFF_INTERVALGENER')= 0 then
        begin
        	InterValEche := InterValEcheDef;
        end else
        begin
        	InterValEche := TOBL.getValue('AFF_INTERVALGENER');
        end;
        //
        if termeEche = 'ECH' then
        begin
        	// Echu
          	TOBL.putValue('AFA_DATEFINFAC',TOBL.getValue('AFA_DATEECHE'));
            if PeriodEche = 'M' then
            begin
            	Date1 := PlusDate(TOBL.getValue('AFA_DATEFINFAC'),InterValEche*(-1),'M');
            	Date1 := PlusDate(Date1,1,'J');
            end else
            begin
            	Date1 := PlusDate(TOBL.getValue('AFA_DATEFINFAC'),-1,'A');
            	Date1 := PlusDate(Date1,1,'J');
            end;
            TOBL.putValue('AFA_DATEDEBUTFAC',Date1 );
        end else
        begin
        	// à Echoir
          	TOBL.putValue('AFA_DATEDEBUTFAC',TOBL.getValue('AFA_DATEECHE'));
            if PeriodEche = 'M' then
            begin
            	Date1 := PlusDate(TOBL.getValue('AFA_DATEDEBUTFAC'),InterValEche,'M');
            	Date1 := PlusDate(Date1,-1,'J');
            end else
            begin
            	Date1 := PlusDate(TOBL.getValue('AFA_DATEDEBUTFAC'),1,'A');
            	Date1 := PlusDate(Date1,-1,'J');
            end;
            TOBL.putValue('AFA_DATEFINFAC', Date1);
        end;
        TOBL.UpdateDB(false);
        PB.Progress := PB.Progress + 1;
        self.refresh;
      end;
    end;
    ferme (QQ);
  FINALLY
  	TOBF.Free;
    Titre.visible := false;
    PB.Visible := false;
    Titre.Refresh;
  END;
end;

end.
