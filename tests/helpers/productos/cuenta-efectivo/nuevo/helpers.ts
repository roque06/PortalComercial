export type { RegistroExcel, LeerRegistrosExcelOpts } from "../../../ceExExcel";
export { leerRegistrosDesdeExcel, marcarCedulaProcesadaEnExcel, marcarCedulasProcesadasEnExcel } from "../../../ceExExcel";

export {
  seleccionarInstrumentoRobusto,
  seleccionarDropdownPorCampo,
  esperarYClickReintentarPaisIdentificacion,
  validarApnfdYSeleccionarNoSiVacio,
  validarGestionDocumentalSiRequerido,
  cargarDocumentoEnGestionDocumental,
  abrirBpmSiVerificacionConoceCliente,
  esperarPortalListoTrasLogin,
} from "../../../ceNewPortalFlow";

export { agregarRelacionadoSiAplica } from "../../../ceExRelacionados";

export {
  cancelarCasoEnBizagiDesdePortal,
  extraerCasoActivoMpn,
  abrirSolicitudCumplimientoEnBizagiDesdePortal,
  abrirSolicitudPlaftEnBizagiDesdePortal,
} from "../../../ceExBizagi";

export { seleccionarProductoCuentaEfectivoNuevo } from "./productSelection";
