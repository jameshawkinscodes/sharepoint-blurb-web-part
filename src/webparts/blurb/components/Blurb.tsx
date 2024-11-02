import * as React from 'react';
import { IBlurbProps } from './IBlurbProps';
import { Icon } from '@fluentui/react/lib/Icon';

export const Blurb: React.FunctionComponent<IBlurbProps> = ({ containers = [], onContainerClick }) => {
  const [containerList, setContainerList] = React.useState(containers);

  const handleContainerClick = (index: number): void => {
    onContainerClick(index);
  };

  React.useEffect(() => {
    setContainerList(containers);
  }, [containers]);

  return (
    <div style={{ textAlign: 'center' }}>
      <div style={{ display: 'flex', justifyContent: 'center', flexWrap: 'wrap' }}>
        {containerList.map((container, index) => (
          <div
            key={index}
            onClick={() => handleContainerClick(index)}
            style={{
              backgroundColor: container.backgroundColor,
              border: `2px solid ${container.borderColor}`,
              borderRadius: container.borderRadius,
              margin: '10px',
              padding: '20px',
              width: '200px',
              cursor: 'pointer',
              color: container.fontColor || '#000000', // Fallback to a default color if not set
            }}
          >
            {container.icon && (
              <Icon
                iconName={container.icon}
                style={{ fontSize: 40, color: container.fontColor || '#000000' }}
                aria-hidden="true"
              />
            )}
            <h3 style={{ color: container.fontColor || '#000000' }}>{container.title || 'Default Title'}</h3>
            <p style={{ color: container.fontColor || '#000000' }}>{container.text || 'Add text'}</p>
          </div>
        ))}
      </div>
    </div>
  );
};
