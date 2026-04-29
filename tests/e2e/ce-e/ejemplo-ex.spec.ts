import { test, expect } from '@playwright/test';

test('test', async ({ page }) => {
    await page.getByText('Descartar').click();
    await page.getByRole('combobox', { name: 'Motivo Coincidencias OFAC:' }).click();
    await page.getByRole('combobox', { name: 'Motivo Coincidencias OFAC:' }).click();
    await page.getByRole('combobox', { name: 'Motivo Coincidencias OFAC:' }).click();
    await page.getByRole('option', { name: 'Coincidencia descartada por no corresponderse con la persona incluida en las Listas de Control Internas de Clientes', exact: true }).click();
    await page.getByRole('button', { name: 'Siguiente' }).click();
    await page.getByRole('button', { name: 'Aceptar' }).click();




});