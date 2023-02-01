import { useEffect, useState } from 'react';
import XLSX from 'xlsx-js-style';

import MainData from './data/data.json';
import { elementData } from './data/element';
import { testSuiteData } from './data/testSuite';
import { projectData } from './data/project';
import {
  columnWidthChanges,
  updateActionsBackgroundStyle,
  updateSingleCellStyle,
  updateStepsBackgroundStyle,
  updateStepsHeaderStyle,
  updateUserHeadersStyle,
  updateUserInfoStyle,
} from './components/styles';

function App() {
  const [data, setData] = useState([]);

  useEffect(() => {
    setData(MainData);
  }, []);

  //todo You can use the option to add photos for a fee.

  const handleClick = () => {
    console.log('data downloaded');

    //* Test Suite
    let testSuite;
    for (let i = 0; i < testSuiteData.length; i++) {
      if (!testSuite || testSuite.id !== testSuiteData[i].id) {
        testSuite = testSuiteData.find(
          (item) => item.id === testSuiteData[i]?.id
        );
      }
    }

    //* Project details
    var upperData = [
      ['SECENORY'],
      [''],
      ['Project Details', ''],
      ['ID', data?.id],
      ['Last Update', new Date(data?.lastUpdated).toLocaleString()],
      ['Members', data?.members?.toString()],
      ['Organization', data?.organization],
      ['Specname', data?.specname],
      ['Tags', data?.tags?.toString()],
      ['Test Data', data?.testdata === null ? '' : data?.testdata],
      ['Test Suite', testSuite?.testsuitename],
      ['Timestamp', new Date(data?.timestamp).toLocaleString()],
      ['Type', data?.type],
    ];

    //* Steps headers
    var lowerData = [];
    lowerData.push(
      [''],
      ['Test Case Steps', ''],
      ['Count', 'Name', 'Action', 'ElementID', 'Project for', 'Text']
    );

    var content = data?.content;

    //* Element Name
    let element;
    let project;
    //* Searching projects into the elementID data file
    for (let i = 0; i < content.length; i++) {
      if (!element || element.id !== content[i].elementID) {
        element = elementData.find(
          (item2) => item2.id === content[i].elementID
        );
      }

      if (!project || project.id !== element?.projectID) {
        project = projectData.find((item) => item.id === element?.projectID);
      }

      lowerData.push([
        i + 1,
        content[i]?.name,
        content[i]?.action,
        element?.name ? element?.name : 'undefined',
        project?.name ? project?.name : 'unknown',
        content[i]?.text === undefined ? '' : content[i]?.text,
      ]);
    }

    //? Creating workbook and worksheet
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet(upperData.concat(lowerData));

    //! STYLES
    //* User headers
    const user = XLSX.utils.decode_range('A4:A13');
    updateUserHeadersStyle(worksheet, user);

    //* User Info
    const userInfo = XLSX.utils.decode_range('B4:B13');
    updateUserInfoStyle(worksheet, userInfo);

    //* Single blocks
    updateSingleCellStyle(worksheet);

    //* Steps header
    const stepsHeader = XLSX.utils.decode_range('A16:F16');
    updateStepsHeaderStyle(worksheet, stepsHeader);

    //* Common background
    const steps = XLSX.utils.decode_range('A17:F33');
    updateStepsBackgroundStyle(worksheet, steps);

    //* Actions background color change
    const actions = XLSX.utils.decode_range('C17:C33');
    updateActionsBackgroundStyle(worksheet, actions);

    //* column width changes
    columnWidthChanges(worksheet);
    //! STYLES END

    //? Creating a sheet and file
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Details1');
    XLSX.writeFile(workbook, 'Details.xlsx');
  };

  return (
    <div className="App">
      <h1>Data</h1>
      <button onClick={handleClick}>EXPORT</button>

      <h2>Project Details</h2>
      <p>
        id - last update - members - organization - specname - tags - test data
        - test suite - timestamp - type
      </p>
      <h2>Steps Details</h2>
      <p>action - elementID - id - name - text - storename</p>
    </div>
  );
}

export default App;
