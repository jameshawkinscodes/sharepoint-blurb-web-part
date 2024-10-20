import * as React from 'react';
import { IBlurbProps } from './IBlurbProps';
import { Icon } from '@fluentui/react/lib/Icon';

export const Blurb: React.FunctionComponent<IBlurbProps> = (props) => {
  // Initialize containers with props.containers or an empty array
  const [containers, setContainers] = React.useState(props.containers || []);

  // Handle when a container is clicked
  const handleContainerClick = (index: number): void => {
    props.onContainerClick(index);  // Notify the web part that a container was clicked
  };

  // Update containers whenever props.containers changes
  React.useEffect(() => {
    if (props.containers) {
      setContainers(props.containers);
    }
  }, [props.containers]);

  return (
    <div style={{ textAlign: 'center' }}>
      <div style={{ display: 'flex', justifyContent: 'center', flexWrap: 'wrap' }}>
        {containers.map((container, index) => (
          <div 
            key={index} 
            onClick={() => handleContainerClick(index)} 
            style={{
              backgroundColor: container.backgroundColor,
              border: `2px solid ${container.borderColor}`,
              margin: '10px',
              padding: '20px',
              width: '200px',
              cursor: 'pointer'
            }}
          >
            {container.icon && (
              <Icon iconName={container.icon} style={{ fontSize: 40 }} aria-hidden="true" />
            )}
            <h3>{container.title || 'No Title'}</h3> {/* Changed 'Test' to 'No Title' */}
            <p>{container.text || ''}</p>
          </div>
        ))}
      </div>
    </div>
  );
};
