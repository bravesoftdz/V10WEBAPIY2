unit QMenuOBJ;

interface

Procedure ObjMenuDispatch (Num : Integer);
Procedure ObjModeMenuDispatch (Num : Integer);

implementation

uses {$IFDEF EAGLCLIENT} Maineagl, {$ELSE EAGLCLIENT} FE_Main, {$ENDIF}
     UFormuleMulBP,
     QUFGBPHPR_TOF, QUFCBPBUDGETPREV_TOF, QUFMBPMUL_TOF, QUFSBPARBREPREV_TOF, QUFSBPBUDGETR3_TOF,
     QUFSLARBRE_TOF, QUFVBPARBREMODIF_TOF, QUFVBPDUPLIQUER_TOF, QBPLOI_TOM, QUFQBPBUDGET_TOF, QBPCALENDREP_TOM,
     QUFVBPCALENDJOURM_TOF, QUFVBPDETAILJOURM_TOF, QUFVBPDETCALEND_TOF,QUFVBPEVOLUTION2_TOF,QUFVBPMULTISOC_TOF,QUFVBPNEWVAL_TOF,
     QUFVBPSUBNIVAUTO_TOF;

Procedure ObjMenuDispatch (Num : Integer);
begin
  Case Num of
  { 216100 : Objectifs }
    216110 : AglLanceFiche('Q', 'QUFMBPSESSIONENT', '', '', ''); { Session }
    //216150 : AglLanceFiche('Q', 'QUFVBPGENCUB', '', '', ''); { Génération }
    216120 : AglLanceFiche('Q', 'QUFCBPBUDGETPREV', '', '', '');  { Cube }
    216130 : AglLanceFiche('Q', 'QUFSBPBUDGETR3', '', '', '');  { Tableau de bord }
    216141 : AglLanceFiche('Q', 'QUFQBPBUDGET', '', '', 'CONTEXTE=PERIODE'); { Editions }
    { 216200 : Prévisions }
  end;
end;

Procedure ObjModeMenuDispatch (Num : Integer);
begin
  Case Num of
  { 216100 : Objectifs }
    216110 : AglLanceFiche('Q', 'QUFMBPSESSIONENT', '', '', '');   { Session }
    216120 : AglLanceFiche('Q', 'QUFCBPBUDGETPREV', '', '', '');   { Cube }
    216130 : AglLanceFiche('Q', 'QUFSBPBUDGETR3', '', '', '');   { Tableau de bord }
    216150 : AglLanceFiche('Q', 'QUFMBPCALENDREP','','','');   { Calendrier }
    216160 : AglLanceFiche('Q', 'QUFQBPBUDGET', '', '', 'CONTEXTE=PERIODE');  {  Restitution }
    216170 : AglLanceFiche('Q', 'QUFSBPOBJBO','','','');  {  Restitution }
  end;
end;

end.

