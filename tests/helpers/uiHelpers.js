export {
    esperarDropdownCargado,
    manejarReintentar,
    seleccionarDropdown,
    seleccionarDefaultSiVacio,
    clickReintentarListaSiVisible,
    seleccionarDropdownConReintentoYReintentarBtn,
} from './ui/dropdowns-basic.js';

export {
    seleccionarDropdownConReintento,
    asegurarIdentificacionHabilitada,
    seleccionarDropdownFiltrableConReintentar,
} from './ui/dropdowns-advanced.js';

export {
    llenarFecha,
    llenarFechaSiVisibleYVacia,
    clickSwitch,
    llenarCampoSiVacio,
    llenarInputNumber,
    llenarInputNumberSiVacio,
    llenarInputMask,
    ejecutarSiLabelVisible,
    llenarCampoPorLabel,
    llenarCampoYEnter,
    clickBotonPorLabel,
    llenarFechaMinimaYDepurar,
} from './ui/inputs.js';

export { fieldContainerByLabel } from './ui/shared.js';

export {
    corregirErroresPrimeVue,
    seleccionarNoSiVacioSelectButton,
    seleccionarNoSiVacioPorPregunta,
    completarSiHayRequeridos,
    validarCorreoPredeterminadoYCorregir,
} from './ui/validation.js';

export {
    capturarCuentaComoPNG,
    unirPNGsEnUnPDF,
} from './ui/evidence.js';
