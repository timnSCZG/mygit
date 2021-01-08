# -*- coding: utf-8 -*-
# Part of Odoo. See LICENSE file for full copyright and licensing details.

import xlwt
import xlrd
import datetime
import unicodedata
import base64
import io
from io import StringIO
import csv
from datetime import datetime
from odoo import api, fields, models, _


class SaleReportOut(models.Model):
    _name = 'sale.report.out'
    _description = 'sale order report'

    sale_data = fields.Char('Name', size=256)
    file_name = fields.Binary('Sale Excel Report', readonly=True)
    sale_work = fields.Char('Name', size=256)
    file_names = fields.Binary(' ', readonly=True)



class WizardWizards(models.Model):
    _name = 'wizard.reports'
    _description = 'sale wizard'

#sale order excel report button actions
#     @api.multi
    def action_sale_report(self):
#XLS report

        custom_value = {}
        label_lists=['SO','SOR','CLIENTREF','BARCODE','DEFAULTCODE','NAME','QTY','VENDORPRODUCTCODE','TITLE', 'PARTNERNAME', 'EMAIL', 'PHONE', 'MOBILE', 'STREET', 'STREET2', 'ZIP', 'CITY', 'COUNTRY']
        order = self.env['sale.order'].browse(self._context.get('active_ids', list()))
        workbook = xlwt.Workbook()
        style0 = xlwt.easyxf('font:height 450,bold True; pattern: pattern solid, fore_colour white;align:horiz center;', num_format_str='#,##0.00')
        style1 = xlwt.easyxf('font: name Times New Roman bold on; pattern: pattern solid, fore_colour black;align: horiz center;', num_format_str='#,##0.00')
        style2 = xlwt.easyxf('font:height 200,bold True;align: horiz center;border: left thin,right thin,top thin,bottom thin', num_format_str='#,##0')
        style3 = xlwt.easyxf('font:height 200,bold True;align: horiz center;border: left thin,right thin,top thin,bottom thin', num_format_str='#,##0.00')
        style4 = xlwt.easyxf('font:bold True;  borders:top double;align: horiz right;', num_format_str='#,##0.00')
        style5 = xlwt.easyxf('font: name Times New Roman bold on;align: horiz center;', num_format_str='#,##0')
        style6 = xlwt.easyxf(' font: name Times New Roman bold on;', num_format_str='#,##0.00',)
        style7 = xlwt.easyxf('font:bold True;  borders:top double;', num_format_str='#,##0.00')
        style8 = xlwt.easyxf('font:bold True;align: horiz center;', num_format_str='#,##0.00')
        style9 = xlwt.easyxf('align: wrap yes,vert centre, horiz left;border: left thin,right thin,top thin,bottom thin')
        style10 = xlwt.easyxf('font:height 160,bold True;align: horiz center;border: left thin,right thin,top thin,bottom thin', num_format_str='#,##0.00')
        #统计输出表表头内容

        sheet1 = workbook.add_sheet('报价汇总及分类表',cell_overwrite_ok=True)
        sheet1.write_merge(0,2,1,32,'设备报价汇总及分类表',style0)
        sheet1.write_merge(3,3,26,32,'货币、单位：人民币、元')
        sheet1.write_merge(4,5, 1,2, '机号或者编号', style3)
        sheet1.write_merge(4,5, 3, 4, '型号规格', style3)
        sheet1.write_merge(4, 5, 5, 6, '台数', style3)
        sheet1.write_merge(4,5,7, 8,  '总机长（米）', style3)
        sheet1.write_merge(4,5,9,10, '总重（KG）', style3)
        sheet1.write_merge(4,5, 11,12,'总价', style3)
        sheet1.write_merge(4,4,13,32,'其中各分项总重及总价',style3)
        sheet1.write_merge(5,5,13,14, '胶带总价', style3)
        sheet1.write_merge(5,5,15,16, '电机总价电动滚筒价', style10)
        sheet1.write_merge(5,5,17,18, '减速器总价', style3)
        sheet1.write_merge(5,5,19,20, '电气总价', style3)
        sheet1.write_merge(5,5,21,22, '外购总价', style3)
        sheet1.write_merge(5,5,23,24, '自制件总重', style3)
        sheet1.write_merge(5,5,25,26, '自制件总价', style3)
        sheet1.write_merge(5,5,27,28, '钢构总重', style3)
        sheet1.write_merge(5,5,29,30, '钢构总价', style3)
        sheet1.write_merge(5,5,31,32, '价格幅度%', style3)
#宽度设置
        col_i = 2
        for col_i in range(33):
            sheet1.col(col_i).width = 256*5
        next
        ts = sjjc = amount_weight = amount_total = 0.0;
        aasubprice1 =aasubprice2 =aasubprice3=aasubprice4=aasubprice5=aasubprice6=aasubprice7=aasubprice8=aasubprice9=aasubprice10=0.0
        m=6;
        for rec in order:
            allsubprice1 = allsubprice2 =allsubprice3=allsubprice4=allsubprice5=allsubprice6=allsubprice7=allsubprice8=allsubprice9=allsubprice10=0.0
            sale = []
            for line in rec.order_line:
                product = {}
                product ['product_id'] = line.product_id.name
                product ['categ_name1'] = line.x_studio_catt1
                product['categ_name'] = line.x_studio_catt
                product ['product_uom'] = line.product_uom.name
                product ['product_uom_qty'] = line.product_uom_qty
                product ['x_studio_odlweight'] = line.x_studio_odlweight
                product ['x_studio_subweight'] = line.x_studio_subweight
                product ['qty_invoiced'] = line.qty_invoiced
                product ['price_unit'] = line.price_unit
                if len(str(line.x_studio_memory))==5:
                    product['x_studio_memory']=" "
                else:
                    product ['x_studio_memory'] = line.x_studio_memory
                product ['price_subtotal'] = line.price_subtotal
                product ['x_studio_odlfactory'] = line.x_studio_odlfactory
                sale.append(product)

            custom_value['products'] = sale
            custom_value ['partner_id'] = rec.partner_id.name
            # custom_value ['partner_ref'] = rec.partner_ref
            # custom_value ['payment_term_id'] = rec.payment_term_id.name
            custom_value ['date_order'] = rec.date_order
            custom_value ['partner_no'] = rec.name
            custom_value ['amount_total'] = rec.amount_total
            custom_value ['amount_weight'] = rec.x_studio_amoweight
            custom_value['x_studio_sbbh'] = rec.x_studio_sbbh
            custom_value['x_studio_wlmc'] = rec.x_studio_wlmc
            custom_value['x_studio_wlrz'] = rec.x_studio_wlrz
            custom_value['x_studio_lidu'] = rec.x_studio_lidu
            custom_value['x_studio_wlwd'] = rec.x_studio_wlwd
            custom_value['x_studio_sbxh'] = rec.x_studio_sbxh
            custom_value['x_studio_sbgg'] = rec.x_studio_sbgg
            custom_value['x_studio_ds'] = rec.x_studio_ds
            custom_value['x_studio_ssl'] = rec.x_studio_ssl
            custom_value['x_studio_hjwd'] = rec.x_studio_hjwd
            custom_value['x_studio_spjc'] = rec.x_studio_spjc
            custom_value['x_studio_sjjc'] = rec.x_studio_sjjc
            custom_value['x_studio_tsgd'] = rec.x_studio_tsgd
            custom_value['x_studio_qjd'] = rec.x_studio_qjd
            custom_value['x_studio_ljfs'] = rec.x_studio_ljfs
            custom_value['x_studio_sjry'] = rec.x_studio_sjry
            custom_value['x_studio_ts'] = rec.x_studio_ts
            custom_value['x_studio_djgl'] = rec.x_studio_djgl
            custom_value['x_studio_djdy'] = rec.x_studio_djdy
            custom_value['x_studio_qdfs'] = rec.x_studio_qdfs
            custom_value['x_pricelist'] = rec.pricelist_id.name

            #新建以报价单为表名的sheet
            sheet = workbook.add_sheet(rec.name)
            sheet.write_merge(0,0,1,16,custom_value['partner_id']+'选型',style2)
            sheet.write_merge(2, 2, 1, 2, '设备编号', style3)
            sheet.write_merge(2, 2, 3, 4, custom_value['x_studio_sbbh'], style3)
            sheet.write_merge(2, 2, 5, 6, '设备型号', style3)
            sheet.write_merge(2, 2, 7, 8, custom_value['x_studio_sbxh'], style3)
            sheet.write_merge(2, 2, 9, 10, '水平机长(米)', style3)
            sheet.write_merge(2, 2, 11, 12, custom_value['x_studio_spjc'], style3)
            sheet.write_merge(2, 2, 13, 14, '设计人员', style3)
            sheet.write_merge(2, 2, 15, 16, custom_value['x_studio_sjry'], style3)
            sheet.write_merge(3, 3, 1, 2, '物料名称', style3)
            sheet.write_merge(3, 3, 3, 4, custom_value['x_studio_wlmc'], style3)
            sheet.write_merge(3, 3, 5, 6, '设备规格(mm)', style3)
            sheet.write_merge(3, 3, 7, 8, custom_value['x_studio_sbgg'], style3)
            sheet.write_merge(3, 3, 9, 10, '实际机长(米)', style3)
            sheet.write_merge(3, 3, 11, 12, custom_value['x_studio_sjjc'], style3)
            sheet.write_merge(3, 3, 13, 14, '台数', style3)
            sheet.write_merge(3, 3, 15, 16, custom_value['x_studio_ts'], style3)
            sheet.write_merge(4, 4, 1, 2, '物料容重t/m³', style3)
            sheet.write_merge(4, 4, 3, 4, custom_value['x_studio_wlrz'], style3)
            sheet.write_merge(4, 4, 5, 6, '带速(米/秒)', style3)
            sheet.write_merge(4, 4, 7, 8, custom_value['x_studio_ds'], style3)
            sheet.write_merge(4, 4, 9, 10, '提升高度(米)', style3)
            sheet.write_merge(4, 4, 11, 12, custom_value['x_studio_tsgd'], style3)
            sheet.write_merge(4, 4, 13, 14, '电机功率(KW)', style3)
            sheet.write_merge(4, 4, 15, 16, custom_value['x_studio_djgl'], style3)
            sheet.write_merge(5, 5, 1, 2, '粒度(mm)', style3)
            sheet.write_merge(5, 5, 3, 4, custom_value['x_studio_lidu'], style3)
            sheet.write_merge(5, 5, 5, 6, '输送量(吨/小时)', style3)
            sheet.write_merge(5, 5, 7, 8, custom_value['x_studio_ssl'], style3)
            sheet.write_merge(5, 5, 9, 10, '倾角(°)', style3)
            sheet.write_merge(5, 5, 11, 12, custom_value['x_studio_qjd'], style3)
            sheet.write_merge(5, 5, 13, 14, '电机电压(V)', style3)
            sheet.write_merge(5, 5, 15, 16, custom_value['x_studio_djdy'], style3)
            sheet.write_merge(6, 6, 1, 2, '物料温度', style3)
            sheet.write_merge(6, 6, 3, 4, custom_value['x_studio_wlwd'], style3)
            sheet.write_merge(6, 6, 5, 6, '环境温度', style3)
            sheet.write_merge(6, 6, 7, 8, custom_value['x_studio_hjwd'], style3)
            sheet.write_merge(6, 6, 9, 10, '拉紧方式', style3)
            sheet.write_merge(6, 6, 11, 12, custom_value['x_studio_ljfs'], style3)
            sheet.write_merge(6, 6, 13, 14, '驱动方式', style3)
            sheet.write_merge(6, 6, 15, 16, custom_value['x_studio_qdfs'], style3)


            sheet.write(10, 1, '序 号', style3)
            sheet.write_merge(10, 10, 2, 3, '名 称', style3)
            sheet.write_merge(10, 10, 4, 5, '分 类', style3)
            sheet.write(10, 6,  '单 位', style3)
            sheet.write(10,7, '数 量', style3)
            sheet.write(10, 8, '单 价', style3)
            sheet.write(10,9, '重 量', style3)
            sheet.write(10, 10, '总 重', style3)
            sheet.write_merge(10, 10, 11, 12, '小 计', style3)
            sheet.write_merge(10, 10, 13, 14,'备 注',style3)
            sheet.write_merge(10, 10, 15, 16, '厂 家', style3)


            n = 11; i = 1;
            for product in custom_value['products']:
                sheet.write(n, 1, i, style2)
                sheet.write_merge(n, n, 2, 3, product['product_id'], style3)
                sheet.write_merge(n, n, 4, 5, product['categ_name1'], style3)
                sheet.write(n, 6, product['product_uom'], style3)
                sheet.write(n, 7, product['product_uom_qty'], style3)
                sheet.write(n, 8, product['price_unit'], style3)
                sheet.write(n, 9, product['x_studio_odlweight'], style3)
                sheet.write(n,10, product['x_studio_subweight'], style3)

                sheet.write_merge(n, n, 11, 12, product['price_subtotal'], style3)
                #判断分类并累加小计价格或重量
                if '胶带' in product['categ_name']:
                    allsubprice1 +=product['price_subtotal']
                if '电机' in product['categ_name']:
                    allsubprice2 +=product['price_subtotal']
                if '减速器' in product['categ_name']:
                    allsubprice3 +=product['price_subtotal']
                if '电气' in product['categ_name']:
                    allsubprice4 +=product['price_subtotal']
                if '外购件' in product['categ_name']:
                    allsubprice5 +=product['price_subtotal']
                if '自制件' in product['categ_name']:
                    allsubprice6 +=product['x_studio_subweight']
                if '自制件' in product['categ_name']:
                    allsubprice7 +=product['price_subtotal']
                if '自制件 / 结构件' in product['categ_name']:
                    allsubprice8 +=product['x_studio_subweight']
                if '自制件 / 结构件' in product['categ_name']:
                    allsubprice9 +=product['price_subtotal']
                #
                # if '自制件' in product['categ_name']:
                #     allsubprice10 +=product['x_studio_subweight']

                sheet.write_merge(n, n, 13, 14, product['x_studio_memory'], style3)

                sheet.write_merge(n, n, 15, 16, product['x_studio_odlfactory'], style3)

                n += 1; i += 1;
                # sheet.write_merge(n+1, n+1, 9, 10, 'Untaxed Amount', style7)
                # sheet.write(n+1, 11, custom_value['amount_untaxed'], style4)
                # sheet.write_merge(n+2, n+2, 9, 10, 'Taxes', style7)
                # sheet.write(n+2, 11, custom_value['amount_tax'], style4)

            sheet.write_merge(n + 2, n + 2, 13, 14, '总 重', style7)
            sheet.write_merge(n+2, n+2,15,16, custom_value['amount_weight'], style4)
            sheet.write_merge(n+3, n+3, 13, 14, '总 价', style7)
            sheet.write_merge(n+3,n+3, 15,16, custom_value['amount_total'], style4)
            #汇总统计表输出内容


            sheet1.write_merge(3,3,1,9,'订货单位：'+custom_value ['partner_id'])

            sheet1.write_merge(m,m, 1,2, custom_value['x_studio_sbbh'], style3)
            sheet1.write_merge(m, m, 3, 4, str(custom_value['x_studio_sbxh']), style3)
            sheet1.write_merge(m,m,5,6,custom_value['x_studio_ts'],style3)
            ts=ts+custom_value['x_studio_ts']
            sheet1.write_merge(m,m,7,8,custom_value['x_studio_sjjc'],style3)
            sjjc=sjjc+(custom_value['x_studio_sjjc'])
            sheet1.write_merge(m,m,9,10,custom_value['amount_weight'],style3)
            amount_weight+=(custom_value['amount_weight'])
            # sheet1.write(m,8,custom_value['amount_total'],style6)
            sheet1.write_merge(m,m,11,12, custom_value['amount_total'], style3)
            amount_total+=float(str(custom_value['amount_total']).replace('¥',''))
            # 显示胶带总价
            sheet1.write_merge(m,m,13,14,allsubprice1, style3)
            aasubprice1+=allsubprice1
            sheet1.write_merge(m, m, 15, 16, allsubprice2, style3)
            aasubprice2 += allsubprice2
            sheet1.write_merge(m, m, 17, 18, allsubprice3, style3)
            aasubprice3 += allsubprice3
            sheet1.write_merge(m, m, 19, 20, allsubprice4, style3)
            aasubprice4 += allsubprice4
            sheet1.write_merge(m, m, 21, 22, allsubprice5, style3)
            aasubprice5 += allsubprice5
            sheet1.write_merge(m, m, 23, 24, allsubprice6, style3)
            aasubprice6 += allsubprice6
            sheet1.write_merge(m, m, 25, 26, allsubprice7, style3)
            aasubprice7 += allsubprice7
            sheet1.write_merge(m, m, 27, 28, allsubprice8, style3)
            aasubprice8 += allsubprice8
            sheet1.write_merge(m, m, 29, 30, allsubprice9, style3)
            aasubprice9 += allsubprice9
            sheet1.write_merge(m, m, 31, 32, allsubprice10, style3)
            aasubprice10 += allsubprice10

            #报价调价情况
            # sheet1.write(m,10,custom_value['x_pricelist'],style3)
            m+=1;
        sheet1.write_merge(m,m,1,2,'',style3)
        sheet1.write_merge(m,m,3,4,'总合计',style3)
        sheet1.write_merge(m,m,5,6,ts,style3)
        sheet1.write_merge(m,m,7,8,sjjc,style3)
        sheet1.write_merge(m,m,9,10,amount_weight,style3)
        sheet1.write_merge(m,m,11,12,amount_total,style3)
        sheet1.write_merge(m,m,13,14,aasubprice1,style3)
        sheet1.write_merge(m,m,15,16,aasubprice2,style3)
        sheet1.write_merge(m,m,17,18,aasubprice3,style3)
        sheet1.write_merge(m,m,19,20,aasubprice4,style3)
        sheet1.write_merge(m,m,21,22,aasubprice5,style3)
        sheet1.write_merge(m,m,23,24,aasubprice6,style3)
        sheet1.write_merge(m,m,25,26,aasubprice7,style3)
        sheet1.write_merge(m,m,27,28,aasubprice8,style3)
        sheet1.write_merge(m,m,29,30,aasubprice9,style3)
        #分类汇总重量
        sheet1.write_merge(m,m,31,32,str(aasubprice10),style3)
        sheet1.write_merge(m+1,m+2,1,26,'注：1.本表只作为内部确定报价分析使用；其中钢构是指桁架、栈桥、支柱、走道、栏杆。\n  2.本计价以设计部门选型明细表计算。')
       #使用的价格表
        sheet1.write_merge(m+4,m+4,1,6,custom_value['x_pricelist'])

        wt = self.env['product.pricelist']
        id_needed = wt.search([('name', '=', custom_value['x_pricelist'])]).id
        new = wt.browse(id_needed)
        list1 = new.id
        #sheet1.write_merge(m+6,m+6,1,6,list1)

        wt1 = self.env['product.pricelist.item']
        id_needed1 = wt1.search([('pricelist_id', '=', list1)]).ids
        new1 = wt1.browse(id_needed1)
        tt = []
        for xxx in new1:
            reco={}
            reco['id']=tt
            reco['categ_id']=xxx.categ_id.name
            reco['percent_price']= xxx.percent_price
            reco['ton_price']=xxx.ton_price
            tt.append(reco)
        next
        q = 1;
        sheet1.write_merge(m+q+5,m+q+5,1,6,"分类名称")
        sheet1.write_merge(m+q+5,m+q+5,7,8,"折扣比例%")
        sheet1.write_merge(m+q+5,m+q+5,9,10,"吨价")
        for reco in reco['id']:
            sheet1.write_merge(m+q+6,m+q+6,1,6,reco['categ_id'])
            sheet1.write_merge(m+q+6,m+q+6,7,8,reco['percent_price'])
            sheet1.write_merge(m + q + 6, m + q + 6, 9, 10, reco['ton_price'])
            q += 1;
        next
#CSV report

        # datas = []
        # for values in order:
        #     for value in values.order_line:
        #         if value.product_id.seller_ids:
        #             item = [
        #                     str(value.order_id.name or ''),
        #                     str(''),
        #                     str(''),
        #                     str(value.product_id.barcode or ''),
        #                     str(value.product_id.default_code or ''),
        #                     str(value.product_id.name or ''),
        #                     str(value.product_uom_qty or ''),
        #                     str(value.product_uom.name or ''),
        #                     str(value.product_id.seller_ids[0].product_code or ''),
        #                     str(value.partner_id.title or ''),
        #                     str(value.partner_id.name or ''),
        #                     str(value.partner_id.email or ''),
        #                     str(value.partner_id.phone or ''),
        #                     str(value.partner_id.mobile or ''),
        #                     str(value.partner_id.street or ''),
        #                     str(value.partner_id.street2 or ''),
        #                     str(value.partner_id.zip or ''),
        #                     str(value.partner_id.city or ''),
        #                     str(value.partner_id.country_id.name or ' '),
        #                     ]
        #             datas.append(item)
        #
        # output = StringIO()
        # label = ';'.join(label_lists)
        # output.write(label)
        # output.write("\n")
        #
        # for data in datas:
        #     record = ';'.join(data)
        #     output.write(record)
        #     output.write("\n")
        # data = base64.b64encode(bytes(output.getvalue(),"utf-8"))
        #

        filename = ('Sale Report'+ '.xls')
        workbook.save(filename)
        fp = open(filename, "rb")
        file_data = fp.read()
        out = base64.encodestring(file_data)
# Files actions         
        attach_vals = {
                'sale_data': 'Sale Report'+ '.xls',
                'file_name': out,
                # 'sale_work':'Sale'+ '.csv',
                # 'file_names':data,
            } 
            
        act_id = self.env['sale.report.out'].create(attach_vals)
        fp.close()
        return {
        'type': 'ir.actions.act_window',
        'res_model': 'sale.report.out',
        'res_id': act_id.id,
        'view_type': 'form',
        'view_mode': 'form',
        'context': self.env.context,
        'target': 'new',
        }
                          
        

 





























