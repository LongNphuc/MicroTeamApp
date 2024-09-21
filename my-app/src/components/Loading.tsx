import * as React from 'react';
import { Spinner, SpinnerSize, Text } from '@fluentui/react';

interface Props {
  style?: React.CSSProperties;
  title?: string;
  description?: string;
  className?: string;
}

export function Loading({ style, title, className, description }: Props) {
  console.log('chay ne');
  return (
    <div style={{ display: 'flex', ...style }} className={className ?? "h-screen"}>
      <div style={{ margin: 'auto', textAlign: 'center' }}>
        {!!title && (
          <Text variant="xLarge">
            {title}
          </Text>
        )}
        <div style={{ marginTop: 20 }}>
          <Spinner size={SpinnerSize.large} />
        </div>
        {!!description && (
          <Text variant="large" style={{ marginTop: 20 }}>
            {description}
          </Text>
        )}
      </div>
    </div>
  );
}
