import { Component, OnInit } from '@angular/core';
import * as ExcelJS from 'exceljs'
import { saveAs } from 'file-saver';

@Component({
  selector: 'app-form',
  templateUrl: './form.component.html',
  styleUrls: ['./form.component.scss'],
})
export class FormComponent  implements OnInit {

  constructor() {}

  contactList: any[] = [
    {
      name: 'Mark',
      email: 'mark@example.com',
      phone: '123-456-7890'
    },
    {
      name: 'Jacob',
      email: 'jacob@example.com',
      phone: '987-654-3210'
    },
    {
      name: 'Victoria',
      email: 'victoria@example.com',
      phone: '555-123-4567'
    }
  ];
  name = "John Doe"
  email = "exampl@gmail.com"
  phone = "1234567890"

  ngOnInit() {}

  saveContact() {
    if ([this.name, this.email, this.phone].includes("")) return

    const newContact = {
      name: this.name,
      email: this.email,
      phone: this.phone
    }
    this.contactList.push(newContact)
    this.name = ''
    this.email = ''
    this.phone = ''
  }

  deleteContact(index: number) {
    if (index >= 0 && index < this.contactList.length) {
      this.contactList.splice(index, 1);
    }
  }

  async exportContactsExcel () {
    // ------------------- Construir archivo excel ----------------------------------------------------
    const workbook = new ExcelJS.Workbook()
    const worksheet = workbook.addWorksheet('Report')

    const borderStyle = {
      top: { style: 'thin', color: { argb: 'aaaaaa' } },
      left: { style: 'thin', color: { argb: 'aaaaaa' } },
      bottom: { style: 'thin', color: { argb: 'aaaaaa' } },
      right: { style: 'thin', color: { argb: 'aaaaaa' } }
    }

    const headerStyle = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'c3f7ef' }
    }

    const columnHeaders = 'Nombre,Email,Telefono'.split(',')
    const headerRow = worksheet.addRow(columnHeaders)
    headerRow.eachCell((cell: any) => { cell.fill = headerStyle })
    headerRow.eachCell((cell: any) => { cell.border = borderStyle })
    headerRow.font = { bold: true }

    this.contactList.forEach((contact: any, index: any) => {
      const row = worksheet.addRow(Object.values(contact))
      row.eachCell((cell: any) => { cell.border = borderStyle })
    })


    columnHeaders.forEach((header, index) => {
      const column = worksheet.getColumn(index + 1)
      const headerLength = header.length
      const columnWidth = Math.max(headerLength + 2, 10)

      column.width = columnWidth

      column.eachCell({ includeEmpty: true }, (cell: any) => {
        cell.alignment = { horizontal: 'left', vertical: 'middle' }
      })
    })

    Object.keys(this.contactList[0]).forEach((header, index) => {
      const column = worksheet.getColumn(index + 1)
      const maxLenght = Math.max(...this.contactList.map((d: any) => String(d[header]).length))
      const columnWidth = Math.max(maxLenght + 2, 10)
      column.width = columnWidth
      column.eachCell({ includeEmpty: true }, (cell: any) => {
        cell.alignment = { horizontal: 'left', vertical: 'middle' }
      })
    })


    const excelBuffer = await workbook.xlsx.writeBuffer()
    // Guarda el archivo usando FileSaver.js
    const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, 'Lista de contactos.xlsx');
  }
}
