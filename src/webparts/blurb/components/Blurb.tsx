import * as React from 'react';
import { IBlurbProps } from './IBlurbProps';
import { IconButton } from '@fluentui/react/lib/Button';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './Blurb.module.scss';

export const Blurb: React.FunctionComponent<IBlurbProps> = (props) => {
  const [selectedBlurbIndex, setSelectedBlurbIndex] = React.useState<number | null>(null);

  const handleContainerClick = (index: number, e: React.MouseEvent): void => {
    if (props.isEditMode) {
      e.preventDefault(); // Prevent link navigation in edit mode
    }
    props.onContainerClick(index);
    setSelectedBlurbIndex(index);
  };

  return (
    <div className={styles.blurbContainer}>
      <div className={styles.containerGrid}>
        {props.containers.map((container, index) => {
          const isSelected = selectedBlurbIndex === index;

          const WrapperElement = container.linkUrl ? 'a' : 'div';
          const wrapperProps = container.linkUrl
            ? {
                href: container.linkUrl,
                target: container.linkTarget || '_self',
                className: styles.clickableBlurb,
                onClick: (e: React.MouseEvent) => handleContainerClick(index, e),
              }
            : {
                onClick: (e: React.MouseEvent) => handleContainerClick(index, e),
              };

          return (
            <WrapperElement
              key={index}
              {...wrapperProps}
              className={`${styles.container} ${isSelected ? styles.selected : ''}`}
              style={{
                backgroundColor: container.backgroundColor,
                border: `1px solid ${isSelected ? '#333' : container.borderColor}`,
                borderRadius: container.borderRadius,
                boxShadow: isSelected ? '0 0 5px rgba(0, 0, 0, 0.3)' : 'none',
                textDecoration: 'none',
              }}
            >
              {isSelected && (
                <div
                  className={`${styles.toolbar} CanvasControlToolbar-item LightTheme`}
                  style={{ display: isSelected ? 'flex' : 'none' }}
                >
                  <IconButton
                    iconProps={{ iconName: 'ChevronUp' }}
                    title="Move Up"
                    ariaLabel="Move Up"
                    onClick={(e) => {
                      e.stopPropagation();
                      props.onMoveClick(index, 'up');
                    }}
                    className={`${styles.toolbarIcon} ToolbarButton`}
                  />
                  <IconButton
                    iconProps={{ iconName: 'ChevronDown' }}
                    title="Move Down"
                    ariaLabel="Move Down"
                    onClick={(e) => {
                      e.stopPropagation();
                      props.onMoveClick(index, 'down');
                    }}
                    className={`${styles.toolbarIcon} ToolbarButton`}
                  />
                  <IconButton
                    iconProps={{ iconName: 'Delete' }}
                    title="Remove"
                    ariaLabel="Remove"
                    onClick={(e) => {
                      e.stopPropagation();
                      props.onRemoveClick(index, props.containers.length - 1);
                    }}
                    className={`${styles.toolbarIcon} ToolbarButton`}
                  />
                </div>
              )}

              {container.icon && (
                <Icon
                  iconName={container.icon}
                  style={{ fontSize: 40, color: container.fontColor }}
                  aria-hidden="true"
                />
              )}
              <h3 style={{ color: container.fontColor }}>{container.title || 'Blurb Title'}</h3>
              <p style={{ color: container.fontColor }}>{container.text || 'Add text'}</p>
            </WrapperElement>
          );
        })}
      </div>
    </div>
  );
};
