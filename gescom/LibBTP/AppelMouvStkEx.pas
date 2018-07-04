unit AppelMouvStkEx;


interface

Uses StdCtrls,Controls,Classes,sysutils,forms,
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      db,FE_main,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
      ComCtrls,Hpanel, Math,HCtrls,HEnt1,HMsgBox,UTOF,Vierge,UTOB,AglInit,LookUp,EntGC,SaisUtil,
			UtofMouvStkEx,uEntCommun;

procedure ConsultMouvEx (cledoc : R_Cledoc;Action : TactionFiche);

implementation

procedure ConsultMouvEx (cledoc : R_Cledoc;Action : TactionFiche);
begin
  GCLanceFiche_MouvStkEx('GC','GCMOUVSTKEX','','',Cledoc.NaturePiece+';'+
                                                 DateToStr(Cledoc.DatePiece)+';'+
                                                 cledoc.Souche+';'+
                                                 IntToStr(cledoc.NumeroPiece)+';'+
                                                 InttoStr(cledoc.Indice)+';ACTION=CONSULTATION');
end;

end.
 
