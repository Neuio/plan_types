import React, { useState, useEffect } from "react";
import { Button, Table, Tooltip, message, Input, InputNumber, Checkbox, Space, Modal, Row, Col, AutoComplete } from "antd";
import { CheckOutlined, CloseOutlined, InfoCircleOutlined, ExclamationCircleOutlined, CopyOutlined, PlusOutlined, FileExcelOutlined } from "@ant-design/icons";
import Papa from "papaparse";
import * as XLSX from 'xlsx';

const CsvTable = () => {
  const [data, setData] = useState([]);
  const [columns, setColumns] = useState([]);
  const [columnDescriptions, setColumnDescriptions] = useState({});
  const [columnUniqueValues, setColumnUniqueValues] = useState({});
  const [columnTypes, setColumnTypes] = useState({});
  const [isModalVisible, setIsModalVisible] = useState(false);
  const [isConfirmModalVisible, setIsConfirmModalVisible] = useState(false);
  const [selectedRecord, setSelectedRecord] = useState(null);
  const [originalRecord, setOriginalRecord] = useState(null);
  const [pendingChanges, setPendingChanges] = useState([]);
  const [allHeaders, setAllHeaders] = useState([]);
  const [allDescriptions, setAllDescriptions] = useState([]);
  const [originalCsv, setOriginalCsv] = useState('');

  const calcColumnWidths = (headersToCalc, rowsToCalc, originalHeaders) => {
    return headersToCalc.map((header) => {
      let maxLen = header.length;
      rowsToCalc.forEach(row => {
        const colIdx = originalHeaders.indexOf(header);
        const val = row[colIdx];
        if (val && val.toString().length > maxLen) {
          maxLen = val.toString().length;
        }
      });
      return Math.min(maxLen * 12 + 20, 400);
    });
  };

  const isBooleanColumn = (columnData) => {
    const uniqueValues = new Set(
      columnData.filter(v => v !== null && v !== undefined && v !== "")
    );
    
    const valuesArray = Array.from(uniqueValues).map(v => v.toString().toLowerCase());
    
    const isTrueFalse = valuesArray.every(v => v === 'true' || v === 'false');
    const isZeroOne = valuesArray.every(v => v === '0' || v === '1');
    
    return isTrueFalse || isZeroOne;
  };

  const isNumericColumn = (columnData) => {
    const nonEmptyValues = columnData.filter(v => v !== null && v !== undefined && v !== "");
    if (nonEmptyValues.length === 0) return false;
    
    return nonEmptyValues.every(v => {
      const str = v.toString().trim();
      return !isNaN(str) && str !== '';
    });
  };

  const convertToBoolean = (value) => {
    if (value === null || value === undefined || value === "") return null;
    const strValue = value.toString().toLowerCase();
    return strValue === 'true' || strValue === '1';
  };

  const renderBooleanCell = (value) => {
    const boolValue = convertToBoolean(value);
    if (boolValue === null) return null;
    
    return boolValue ? (
      <CheckOutlined style={{ color: '#52c41a', fontSize: 18 }} />
    ) : (
      <CloseOutlined style={{ color: '#ff4d4f', fontSize: 18 }} />
    );
  };

  const CustomFilterDropdown = ({ dataIndex, uniqueValues, setSelectedKeys, selectedKeys, confirm, clearFilters }) => {
    const [searchValue, setSearchValue] = useState('');

    const filteredValues = uniqueValues.filter(item =>
      item.value.toString().toLowerCase().includes(searchValue.toLowerCase())
    );

    useEffect(() => {
      if (searchValue) {
        const allFilteredValues = filteredValues.map(item => item.value);
        setSelectedKeys(allFilteredValues);
      }
    }, [searchValue]);

    const handleCheckboxChange = (value, checked) => {
      const newKeys = checked
        ? [...selectedKeys, value]
        : selectedKeys.filter(k => k !== value);
      setSelectedKeys(newKeys);
    };

    const handleSelectAll = () => {
      const allValues = filteredValues.map(item => item.value);
      setSelectedKeys(allValues);
    };

    const handleDeselectAll = () => {
      setSelectedKeys([]);
    };

    return (
      <div style={{ padding: 8, width: 250 }}>
        <Input
          placeholder={`Поиск...`}
          value={searchValue}
          onChange={e => setSearchValue(e.target.value)}
          style={{ marginBottom: 8, display: 'block' }}
        />
        <div style={{ marginBottom: 8 }}>
          <Button size="small" onClick={handleSelectAll} style={{ marginRight: 4 }}>
            Выбрать все
          </Button>
          <Button size="small" onClick={handleDeselectAll}>
            Снять все
          </Button>
        </div>
        <div style={{ maxHeight: 200, overflowY: 'auto', marginBottom: 8 }}>
          {filteredValues.map(item => (
            <div key={item.value} style={{ padding: '4px 0' }}>
              <Checkbox
                checked={selectedKeys.includes(item.value)}
                onChange={e => handleCheckboxChange(item.value, e.target.checked)}
              >
                {item.text}
              </Checkbox>
            </div>
          ))}
        </div>
        <Space>
          <Button
            type="primary"
            size="small"
            onClick={() => confirm()}
            style={{ width: 90 }}
          >
            OK
          </Button>
          <Button
            size="small"
            onClick={() => {
              clearFilters();
              setSearchValue('');
            }}
            style={{ width: 90 }}
          >
            Сброс
          </Button>
        </Space>
      </div>
    );
  };

  const handleRowClick = (record) => {
    setSelectedRecord({ ...record });
    setOriginalRecord({ ...record });
    setIsModalVisible(true);
  };

  const handleModalClose = () => {
    setIsModalVisible(false);
    setSelectedRecord(null);
    setOriginalRecord(null);
  };

  const handleAddNewRow = () => {
    if (allHeaders.length === 0) {
      message.warning("Сначала загрузите данные");
      return;
    }

    const newRow = {};
    allHeaders.forEach(header => {
      newRow[header] = '';
    });
    newRow.key = data.length;

    setSelectedRecord(newRow);
    setOriginalRecord(null);
    setIsModalVisible(true);
  };

  const handleCopyRow = () => {
    if (!selectedRecord) return;

    const copiedRow = { ...selectedRecord };
    delete copiedRow.key;
    copiedRow.key = data.length;

    handleModalClose();

    setTimeout(() => {
      setSelectedRecord(copiedRow);
      setOriginalRecord(null);
      setIsModalVisible(true);
    }, 100);
  };

  const getChanges = () => {
    if (!selectedRecord || !originalRecord) return [];
    
    const changes = [];
    Object.keys(selectedRecord).forEach(key => {
      if (key !== 'key' && selectedRecord[key] !== originalRecord[key]) {
        changes.push({
          field: key,
          oldValue: originalRecord[key],
          newValue: selectedRecord[key]
        });
      }
    });
    return changes;
  };

  const handleSave = () => {
    if (!originalRecord) {
      const updatedData = [...data, selectedRecord];
      setData(updatedData);
      
      // Пересоздаём originalCsv с новой строкой
      const headers = allHeaders;
      const csvRows = [
        headers.join(';'),
        allDescriptions.join(';'),
        ...updatedData.map(row => headers.map(header => row[header] || '').join(';'))
      ];
      setOriginalCsv(csvRows.join('\n'));
      
      message.success("Новая строка добавлена!");
      handleModalClose();
      return;
    }

    const changes = getChanges();
    
    if (changes.length === 0) {
      message.info("Изменений не обнаружено");
      handleModalClose();
      return;
    }

    setPendingChanges(changes);
    setIsConfirmModalVisible(true);
  };

  const handleConfirmSave = () => {
    const updatedData = data.map(item => 
      item.key === selectedRecord.key ? selectedRecord : item
    );
    setData(updatedData);
    
    // Пересоздаём originalCsv с обновлёнными данными
    const headers = allHeaders;
    const csvRows = [
      headers.join(';'),
      allDescriptions.join(';'),
      ...updatedData.map(row => headers.map(header => row[header] || '').join(';'))
    ];
    setOriginalCsv(csvRows.join('\n'));
    
    message.success("Изменения сохранены!");
    setIsConfirmModalVisible(false);
    handleModalClose();
  };

  const handleCancelConfirm = () => {
    setIsConfirmModalVisible(false);
  };

  const handleExportToExcel = () => {
    if (data.length === 0 || allHeaders.length === 0) {
      message.warning("Нет данных для экспорта");
      return;
    }

    try {
      const excelData = [
        allHeaders,
        allDescriptions,
        ...data.map(row => allHeaders.map(header => row[header] || ''))
      ];

      const ws = XLSX.utils.aoa_to_sheet(excelData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Data");

      XLSX.writeFile(wb, `data_${new Date().toISOString().split('T')[0]}.xlsx`);
      message.success("Файл успешно сохранён!");
    } catch (err) {
      console.error(err);
      message.error("Ошибка при экспорте в Excel");
    }
  };

  const handleCopyToClipboard = async () => {
    if (!originalCsv || data.length === 0) {
      message.warning("Нет данных для копирования");
      return;
    }

    try {
      const originalLines = originalCsv.split('\n');
      const headers = Papa.parse(originalLines[0], { delimiter: ';' }).data[0];
      
      const resultLines = [
        originalLines[0],
        originalLines[1]
      ];
      
      data.forEach((row, index) => {
        const originalLineIndex = index + 2;
        
        if (originalLineIndex < originalLines.length) {
          const originalLine = originalLines[originalLineIndex];
          const parsed = Papa.parse(originalLine, { delimiter: ';' });
          const originalValues = parsed.data[0];
          
          let hasChanges = false;
          for (let i = 0; i < headers.length; i++) {
            const header = headers[i];
            const originalValue = originalValues[i] || '';
            const currentValue = row[header] !== undefined ? row[header] : '';
            
            if (originalValue !== currentValue) {
              hasChanges = true;
              break;
            }
          }
          
          if (hasChanges) {
            const values = headers.map(header => row[header] !== undefined ? row[header] : '');
            resultLines.push(values.join(';'));
          } else {
            resultLines.push(originalLine);
          }
        } else {
          const values = headers.map(header => row[header] !== undefined ? row[header] : '');
          resultLines.push(values.join(';'));
        }
      });
      
      const result = resultLines.join('\n');
      
      await navigator.clipboard.writeText(result);
      message.success("Данные скопированы в буфер обмена!");
    } catch (err) {
      console.error(err);
      message.error("Ошибка при копировании в буфер обмена");
    }
  };

  const handlePaste = async () => {
    try {
      const text = await navigator.clipboard.readText();
      if (!text) {
        message.warning("Буфер обмена пуст");
        return;
      }

      setOriginalCsv(text);

      const parsed = Papa.parse(text.trim());
      const rows = parsed.data;

      if (rows.length < 3) {
        message.warning("Недостаточно строк в CSV (нужно минимум 3)");
        return;
      }

      const headers = rows[0];
      const descriptions = rows[1];
      const dataRows = rows.slice(2);

      setAllHeaders(headers);
      setAllDescriptions(descriptions);

      const validHeaders = headers.filter((header, colIdx) => {
        return dataRows.some(row => {
          const val = row[colIdx];
          return val !== null && val !== undefined && val !== "";
        });
      });

      const validDescriptions = validHeaders.map(hdr => descriptions[headers.indexOf(hdr)]);

      const descriptionsMap = {};
      validHeaders.forEach((header, idx) => {
        descriptionsMap[header] = validDescriptions[idx] || "";
      });
      setColumnDescriptions(descriptionsMap);

      const formattedData = dataRows.map((row, idx) => {
        const obj = {};
        headers.forEach((header, colIdx) => {
          obj[header] = row[colIdx] !== undefined ? row[colIdx] : '';
        });
        obj.key = idx;
        return obj;
      });

      const uniqueValuesMap = {};
      const typesMap = {};
      
      validHeaders.forEach(header => {
        const columnData = formattedData.map(item => item[header]);
        const uniqueValues = Array.from(
          new Set(columnData.filter(v => v !== null && v !== undefined && v !== ""))
        ).sort((a, b) => a.toString().localeCompare(b.toString()));
        
        uniqueValuesMap[header] = uniqueValues;
        
        if (isBooleanColumn(columnData)) {
          typesMap[header] = 'boolean';
        } else if (isNumericColumn(columnData)) {
          typesMap[header] = 'numeric';
        } else {
          typesMap[header] = 'text';
        }
      });
      
      setColumnUniqueValues(uniqueValuesMap);
      setColumnTypes(typesMap);

      const widths = calcColumnWidths(validHeaders, dataRows, headers);

      const cols = validHeaders.map((header, idx) => {
        const columnData = formattedData.map(item => item[header]);
        const isBool = isBooleanColumn(columnData);

        const uniqueValues = Array.from(
          new Set(columnData.filter(v => v !== null && v !== undefined && v !== ""))
        )
          .sort((a, b) => a.toString().localeCompare(b.toString()))
          .slice(0, 500)
          .map(value => ({ text: value.toString(), value }));

        return {
          title: (
            <Tooltip title={validDescriptions[idx] || ""}>
              <span>{header}</span>
            </Tooltip>
          ),
          dataIndex: header,
          key: header,
          width: widths[idx],
          filterDropdown: (props) => (
            <CustomFilterDropdown 
              dataIndex={header} 
              uniqueValues={uniqueValues} 
              {...props} 
            />
          ),
          onFilter: (value, record) => record[header] === value,
          render: isBool ? renderBooleanCell : (text => text),
          onCell: () => ({
            style: {
              maxWidth: widths[idx],
              overflow: "hidden",
              whiteSpace: "nowrap",
              textOverflow: "ellipsis",
              textAlign: isBool ? 'center' : 'left',
            },
          }),
        };
      });

      setColumns(cols);
      setData(formattedData);
      message.success("Данные вставлены и отображены!");
    } catch (err) {
      console.error(err);
      message.error("Ошибка при чтении из буфера или обработке CSV");
    }
  };

  return (
    <div style={{ padding: 20 }}>
      <Space>
        <Button type="primary" onClick={handlePaste}>
          Вставить из буфера обмена
        </Button>
        <Button 
          type="default" 
          icon={<CopyOutlined />} 
          onClick={handleCopyToClipboard}
          disabled={data.length === 0}
        >
          Сохранить в буфер обмена
        </Button>
        <Button 
          type="default" 
          icon={<FileExcelOutlined />} 
          onClick={handleExportToExcel}
          disabled={data.length === 0}
        >
          Сохранить в Excel
        </Button>
        <Button 
          type="default" 
          icon={<PlusOutlined />} 
          onClick={handleAddNewRow}
          disabled={allHeaders.length === 0}
        >
          Добавить строку
        </Button>
      </Space>
      
      {data.length > 0 && (
        <Table
          style={{ marginTop: 24 }}
          columns={columns}
          dataSource={data}
          pagination={{
            pageSize: 15,
            showSizeChanger: false,
          }}
          scroll={{ y: window.innerHeight * 0.7, x: "max-content" }}
          rowClassName={() => "fixed-height-row"}
          onRow={(record) => ({
            onClick: () => handleRowClick(record),
            style: { cursor: 'pointer' }
          })}
        />
      )}

      <Modal
        title="Редактирование строки"
        open={isModalVisible}
        onOk={handleSave}
        onCancel={handleModalClose}
        okText="Сохранить"
        cancelText="Закрыть"
        width={900}
        bodyStyle={{ maxHeight: '70vh', overflowY: 'auto', overflowX: 'hidden' }}
        footer={[
          <Button key="copy" icon={<CopyOutlined />} onClick={handleCopyRow}>
            Копировать строку
          </Button>,
          <Button key="cancel" onClick={handleModalClose}>
            Закрыть
          </Button>,
          <Button key="save" type="primary" onClick={handleSave}>
            Сохранить
          </Button>
        ]}
      >
        {selectedRecord && (
          <Row gutter={[16, 16]}>
            {Object.keys(selectedRecord).map((key) => {
              if (key === 'key') return null;
              
              const uniqueValues = columnUniqueValues[key] || [];
              const shouldUseAutoComplete = uniqueValues.length > 0 && uniqueValues.length < 10;
              const columnType = columnTypes[key] || 'text';
              
              return (
                <Col span={12} key={key}>
                  <div style={{ marginBottom: 8 }}>
                    <label style={{ display: 'block', marginBottom: 4, fontWeight: 'bold', wordBreak: 'break-word' }}>
                      {key}
                      {columnDescriptions[key] && (
                        <Tooltip title={columnDescriptions[key]}>
                          <InfoCircleOutlined 
                            style={{ 
                              marginLeft: 6, 
                              color: '#8c8c8c', 
                              fontSize: 14,
                              cursor: 'help'
                            }} 
                          />
                        </Tooltip>
                      )}
                    </label>
                    {shouldUseAutoComplete && columnType !== 'boolean' ? (
                      <AutoComplete
                        style={{ width: '100%' }}
                        value={selectedRecord[key]}
                        onChange={(value) => {
                          if (columnType === 'numeric' && value !== '' && isNaN(value)) {
                            message.warning('Это поле принимает только числовые значения');
                            return;
                          }
                          setSelectedRecord({
                            ...selectedRecord,
                            [key]: value
                          });
                        }}
                        options={uniqueValues.map(v => ({ value: v.toString() }))}
                        placeholder="Введите или выберите..."
                        allowClear
                        filterOption={(inputValue, option) =>
                          option.value.toLowerCase().includes(inputValue.toLowerCase())
                        }
                      />
                    ) : shouldUseAutoComplete && columnType === 'boolean' ? (
                      <AutoComplete
                        style={{ width: '100%' }}
                        value={selectedRecord[key]}
                        onChange={(value) => {
                          setSelectedRecord({
                            ...selectedRecord,
                            [key]: value
                          });
                        }}
                        options={uniqueValues.map(v => ({ value: v.toString() }))}
                        placeholder="Выберите значение..."
                        allowClear
                      />
                    ) : columnType === 'numeric' ? (
                      <InputNumber
                        value={selectedRecord[key] !== '' ? Number(selectedRecord[key]) : null}
                        onChange={(value) => {
                          setSelectedRecord({
                            ...selectedRecord,
                            [key]: value !== null ? value.toString() : ''
                          });
                        }}
                        style={{ width: '100%' }}
                      />
                    ) : (
                      <Input
                        value={selectedRecord[key]}
                        onChange={(e) => {
                          setSelectedRecord({
                            ...selectedRecord,
                            [key]: e.target.value
                          });
                        }}
                        allowClear
                        style={{ width: '100%' }}
                      />
                    )}
                  </div>
                </Col>
              );
            })}
          </Row>
        )}
      </Modal>

      <Modal
        title={
          <div>
            <ExclamationCircleOutlined style={{ color: '#faad14', marginRight: 8 }} />
            Подтвердите изменения
          </div>
        }
        open={isConfirmModalVisible}
        onOk={handleConfirmSave}
        onCancel={handleCancelConfirm}
        okText="Да"
        cancelText="Нет"
        width={600}
      >
        <div>
          <p style={{ marginBottom: 16 }}>Следующие поля будут изменены:</p>
          <div style={{ maxHeight: 400, overflowY: 'auto' }}>
            {pendingChanges.map((change, idx) => (
              <div key={idx} style={{ marginBottom: 12, padding: 8, background: '#f5f5f5', borderRadius: 4 }}>
                <strong>{change.field}:</strong>
                <div style={{ marginTop: 4 }}>
                  <span style={{ color: '#ff4d4f' }}>Было: {change.oldValue || '(пусто)'}</span>
                  {' → '}
                  <span style={{ color: '#52c41a' }}>Станет: {change.newValue || '(пусто)'}</span>
                </div>
              </div>
            ))}
          </div>
        </div>
      </Modal>

      <style>
        {`
          .ant-table-thead > tr > th {
            cursor: default !important;
          }
          .fixed-height-row td {
            padding-top: 4px !important;
            padding-bottom: 4px !important;
          }
          .ant-table-cell {
            border-right: 1px solid #f0f0f0 !important;
          }
          .ant-table-row > td:last-child {
            border-right: none !important;
          }
        `}
      </style>
    </div>
  );
};

export default CsvTable;
