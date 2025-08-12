import * as React from 'react';
import { useRef } from 'react';
import { Stack, PrimaryButton, Text } from '@fluentui/react';

export interface UploadLauncherButtonProps {
	allowMultiple: boolean;
	buttonLabel?: string; // e.g., "Upload file" / "Upload files"
	hint?: string; // e.g., "Drop here to upload"
	disabled?: boolean;
	onFilesChosen: (files: File[]) => void;
}

export const UploadLauncherButton: React.FC<UploadLauncherButtonProps> = ({
	allowMultiple,
	buttonLabel = allowMultiple ? 'Upload files' : 'Upload file',
	hint = allowMultiple ? 'Drop files here' : 'Drop a file here',
	disabled,
	onFilesChosen,
}) => {
	const inputRef = useRef<HTMLInputElement>(null);

	const handleInput = (e: React.ChangeEvent<HTMLInputElement>) => {
		const files = Array.from(e.target.files || []);
		if (files.length) onFilesChosen(files);
		e.currentTarget.value = '';
	};

	const onDrop = (e: React.DragEvent<HTMLDivElement>) => {
		e.preventDefault();
		if (disabled) return;
		if (e.dataTransfer?.files?.length) {
			onFilesChosen(Array.from(e.dataTransfer.files));
		}
	};

	return (
		<Stack
			onDrop={onDrop}
			onDragOver={(e) => e.preventDefault()}
			tokens={{ childrenGap: 8 }}
			styles={{
				root: {
					border: '1px dashed #c8c6c4',
					borderRadius: 10,
					padding: 16,
					background: disabled ? '#f3f2f1' : '#fff',
					textAlign: 'center',
				},
			}}
		>
			<input
				ref={inputRef}
				type="file"
				multiple={allowMultiple}
				onChange={handleInput}
				style={{ display: 'none' }}
				disabled={disabled}
			/>
			<PrimaryButton
				text={buttonLabel}
				onClick={() => inputRef.current?.click()}
				disabled={disabled}
			/>
			<Text variant="small" styles={{ root: { color: '#605e5c' } }}>
				{hint}
			</Text>
		</Stack>
	);
};
