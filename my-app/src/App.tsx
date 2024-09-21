import React from 'react';
import { FluentProvider, teamsLightTheme } from '@fluentui/react-components';
import { Calendar } from '../src/components/Calendar'; // Đường dẫn đến component Calendar


function App() {
  return (
    <FluentProvider theme={teamsLightTheme}>
      <Calendar />
    </FluentProvider>
  );
}

export default App;
