import { Injectable } from '@nestjs/common';
import * as ExcelJS from 'exceljs';
import { faker } from '@faker-js/faker';

@Injectable()
export class AppService {
  async generateExcel(): Promise<Buffer> {
    const workbook = new ExcelJS.Workbook();

    // Crear hojas de trabajo
    const sheetPersonas = workbook.addWorksheet('Personas');
    const sheetDirecciones = workbook.addWorksheet('Direcciones');
    const sheetAutos = workbook.addWorksheet('Autos');

    // Definir títulos para cada hoja
    const personasHeader = ['ID', 'Nombre', 'Edad'];
    const direccionesHeader = ['ID', 'Calle', 'Ciudad', 'País'];
    const autosHeader = ['ID', 'Marca', 'Modelo', 'Año'];

    // Añadir títulos a las hojas
    sheetPersonas.addRow(personasHeader);
    sheetDirecciones.addRow(direccionesHeader);
    sheetAutos.addRow(autosHeader);

    // Función para aplicar estilos a los títulos
    function styleHeader(sheet: any) {
      sheet.getRow(1).eachCell((cell: any) => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FF00FF00' }, // Fondo verde
        };
        cell.font = {
          color: { argb: 'FFFFFFFF' }, // Letras blancas
          bold: true,
        };
      });
    }

    // Aplicar estilos a los títulos en cada hoja
    styleHeader(sheetPersonas);
    styleHeader(sheetDirecciones);
    styleHeader(sheetAutos);

    // Añadir datos aleatorios a la hoja de personas
    for (let i = 0; i < 10; i++) {
      sheetPersonas.addRow([
        i + 1,
        faker.person.fullName(),
        faker.datatype.number({ min: 18, max: 80 }),
      ]);
    }

    // Añadir datos aleatorios a la hoja de direcciones
    for (let i = 0; i < 10; i++) {
      sheetDirecciones.addRow([
        i + 1,
        faker.address.streetAddress(),
        faker.address.city(),
        faker.address.country(),
      ]);
    }

    // Añadir datos aleatorios a la hoja de autos
    for (let i = 0; i < 10; i++) {
      sheetAutos.addRow([
        i + 1,
        faker.vehicle.manufacturer(),
        faker.vehicle.model(),
        faker.datatype.number({ min: 1990, max: 2022 }),
      ]);
    }

    // Guardar el archivo en un buffer
    const arrayBuffer = await workbook.xlsx.writeBuffer();
    const buffer = Buffer.from(arrayBuffer);
    return buffer;
  }
}
