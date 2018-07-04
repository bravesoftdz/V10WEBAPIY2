unit UtilParc;

interface
uses UTOB,SysUtils,Hent1;

procedure ConstitueEltParcFromPere(TOBpere, TOBDetail, TOBA: TOB; Qte: double; Indice : integer);

implementation


procedure ConstitueEltParcFromPere(TOBpere, TOBDetail, TOBA: TOB; Qte: double; Indice : integer);
var N1,N2,N3,N4,N5 : integer;
begin
	if TOBPere.GetValue('BP1_N4')>0 then
  begin
    N1 := TOBPere.GetValue('BP1_N1');
    N2 := TOBPere.GetValue('BP1_N2');
    N3 := TOBPere.GetValue('BP1_N3');
    N4 := TOBPere.GetValue('BP1_N4');
    N5 := 1+Indice;
  end else if TOBPere.GetValue('BP1_N3')>0 then
  begin
    N1 := TOBPere.GetValue('BP1_N1');
    N2 := TOBPere.GetValue('BP1_N2');
    N3 := TOBPere.GetValue('BP1_N3');
    N4 := 1+indice;
    N5 := 0;
  end else if TOBPere.GetValue('BP1_N2')>0 then
  begin
    N1 := TOBPere.GetValue('BP1_N1');
    N2 := TOBPere.GetValue('BP1_N2');
    N3 := 1+indice;
    N4 := 0;
    N5 := 0;
  end else if TOBPere.GetValue('BP1_N1')>0 then
  begin
    N1 := TOBPere.GetValue('BP1_N1');
    N2 := 1+Indice;
    N3 := 0;
    N4 := 0;
    N5 := 0;
  end;
  TOBDetail.putValue('BP1_TIERS',TOBpere.GetValue('BP1_TIERS'));
  TOBDetail.putValue('BP1_N1',N1);
  TOBDetail.putValue('BP1_N2',N2);
  TOBDetail.putValue('BP1_N3',N3);
  TOBDetail.putValue('BP1_N4',N4);
  TOBDetail.putValue('BP1_N5',N5);
  TOBDetail.putValue('BP1_TYPEARTICLE',TOBA.GetValue('GA_TYPEARTICLE'));
  TOBDetail.putValue('BP1_ARTICLE',TOBA.GetValue('GA_ARTICLE'));
  TOBDetail.putValue('BP1_CODEARTICLE',TOBA.GetValue('GA_CODEARTICLE'));
  TOBDetail.putValue('BP1_CODEINTERNE',TOBA.GetValue('BCP_CODEINTERNE'));
  TOBDetail.putValue('BP1_ETATPARC','ES');
  TOBDetail.putValue('BP1_REFFABRICANT',TOBA.GetValue('BP1_REFFABRICANT'));
  if TOBA.GEtValue('BCP_QTEZERO')='X' then
  begin
  	TOBDetail.putValue('BP1_QTE',0);
  end else
  begin
  	TOBDetail.putValue('BP1_QTE',StrToInt(FloatToStr(Qte)));
  end;
  TOBDetail.putValue('BP1_DATEACHAT',TOBPere.GetValue('BP1_DATEACHAT'));
  TOBDetail.putValue('BP1_DATEFINGAR', PlusDate(TOBPere.GetValue('BP1_DATEACHAT'),TOBA.GetValue('BCP_NBMOISGARANTIE'),'M' ));
end;

end.
 